pageextension 50100 ItemListExt extends "Item List"
{
    actions
    {
        addafter(AdjustInventory)
        {
            action(ImportFileFromSharePoint)
            {
                Caption = 'Import File From SharePoint';
                ApplicationArea = All;
                Promoted = true;
                PromotedCategory = Process;
                PromotedIsBig = true;
                Image = GetActionMessages;

                trigger OnAction()
                var
                    SharePointHandler: Codeunit SharePointHandler;
                begin
                    SharePointHandler.Run();
                end;
            }
        }
    }
}

codeunit 50120 SharePointHandler
{
    trigger OnRun()
    begin
        ImportFilesFromSharePoint();
    end;

    var
        CSVBuffer: Record "CSV Buffer" temporary;

    procedure ImportFilesFromSharePoint()
    var
        HttpClient: HttpClient;
        HttpRequestMessage: HttpRequestMessage;
        HttpResponseMessage: HttpResponseMessage;
        Headers: HttpHeaders;
        JsonResponse: JsonObject;
        JsonArray: JsonArray;
        JsonToken: JsonToken;
        JsonTokenLoop: JsonToken;
        JsonValue: JsonValue;
        JsonObjectLoop: JsonObject;
        AuthToken: SecretText;
        SharePointFolderUrl: Text;
        ResponseText: Text;
        FileName: Text;
        FileContent: InStream;
        FileUrl: Text;
        ItemId: Text;
    begin
        // Get OAuth token
        AuthToken := GetOAuthToken();

        if AuthToken.IsEmpty() then
            Error('Failed to obtain access token.');

        // Define the file name
        ItemId := '01AA2EHNNW32ULFEPXI5DIS23SD4S2CEJK';
        // Define the file URL
        FileUrl := 'https://graph.microsoft.com/v1.0/sites/5b3b7cec-cbfe-4893-a638-c18a34c6a394/drive/items/' + ItemId + '/content';

        // Initialize the HTTP request
        HttpRequestMessage.SetRequestUri(FileUrl);
        HttpRequestMessage.Method := 'GET';
        HttpRequestMessage.GetHeaders(Headers);
        Headers.Add('Authorization', SecretStrSubstNo('Bearer %1', AuthToken));

        // Send the HTTP request
        if HttpClient.Send(HttpRequestMessage, HttpResponseMessage) then begin
            if HttpResponseMessage.IsSuccessStatusCode() then begin
                HttpResponseMessage.Content.ReadAs(FileContent);
                // Process the file content (e.g., import into a table)
                ImportFileContent(FileContent);
            end else
                Error('Failed to download file: %1', HttpResponseMessage.HttpStatusCode());
        end else
            Error('Failed to send HTTP request to download file');
    end;

    procedure GetOAuthToken() AuthToken: SecretText
    var
        ClientID: Text;
        ClientSecret: Text;
        TenantID: Text;
        AccessTokenURL: Text;
        OAuth2: Codeunit OAuth2;
        Scopes: List of [Text];
    begin
        ClientID := 'b4fe1687-f1ab-4bfa-b494-0e2236ed50bd';
        ClientSecret := 'huL8Q~edsQZ4pwyxka3f7.WUkoKNcPuqlOXv0bww';
        TenantID := '7e47da45-7f7d-448a-bd3d-1f4aa2ec8f62';
        AccessTokenURL := 'https://login.microsoftonline.com/' + TenantID + '/oauth2/v2.0/token';
        Scopes.Add('https://graph.microsoft.com/.default');
        if not OAuth2.AcquireTokenWithClientCredentials(ClientID, ClientSecret, AccessTokenURL, '', Scopes, AuthToken) then
            Error('Failed to get access token from response\%1', GetLastErrorText());
    end;

    procedure ImportFileContent(FileContent: InStream)
    var
        Item: Record Item;
        LineNo: Integer;
    begin
        CSVBuffer.Reset();
        CSVBuffer.DeleteAll();
        CSVBuffer.LoadDataFromStream(FileContent, ',');
        for LineNo := 2 to CSVBuffer.GetNumberOfLines() do begin
            Item.Init();
            Item.Validate("No.", GetValueAtCell(LineNo, 1));
            Item.Insert(true);
            Item.Validate(Description, GetValueAtCell(LineNo, 2));
            case GetValueAtCell(LineNo, 3) of
                'Inventory':
                    Item.Validate(Type, Item.Type::"Inventory");
                'Service':
                    Item.Validate(Type, Item.Type::"Service");
                'Non-Inventory':
                    Item.Validate(Type, Item.Type::"Non-Inventory");
            end;
            Evaluate(Item.GTIN, GetValueAtCell(LineNo, 4));
            Evaluate(Item."Unit Price", GetValueAtCell(LineNo, 5));
            Item.Validate("Base Unit of Measure", GetValueAtCell(LineNo, 6));
            Item.Modify(true);
        end;
    end;

    local procedure GetValueAtCell(RowNo: Integer; ColNo: Integer): Text
    begin
        if CSVBuffer.Get(RowNo, ColNo) then
            exit(CSVBuffer.Value)
        else
            exit('');
    end;
}
