pageextension 50100 CustomerListExt extends "Customer List"
{
    actions
    {
        addafter("Sent Emails")
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
        ImportFilesFromSharePointHandler();
    end;

    procedure ImportFilesFromSharePointHandler()
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
    begin
        // Get OAuth token
        AuthToken := GetOAuthToken();

        if AuthToken.IsEmpty() then
            Error('Failed to obtain access token.');

        // Define the SharePoint folder URL

        SharePointFolderUrl := 'https://graph.microsoft.com/v1.0/sites/5b3b7cec-cbfe-4893-a638-c18a34c6a394/drives/b!7Hw7W_7Lk0imOMGKNMajlK0n-8Wdev9FmPdhx03j5o95rz4xvtmtTIUW5qUH7Jww/root:/Business%20Central:/children';
        // Initialize the HTTP request
        HttpRequestMessage.SetRequestUri(SharePointFolderUrl);
        HttpRequestMessage.Method := 'GET';
        HttpRequestMessage.GetHeaders(Headers);
        Headers.Add('Authorization', SecretStrSubstNo('Bearer %1', AuthToken));

        // Send the HTTP request
        if HttpClient.Send(HttpRequestMessage, HttpResponseMessage) then begin
            // Log the status code for debugging
            //Message('HTTP Status Code: %1', HttpResponseMessage.HttpStatusCode());

            if HttpResponseMessage.IsSuccessStatusCode() then begin
                HttpResponseMessage.Content.ReadAs(ResponseText);
                JsonResponse.ReadFrom(ResponseText);

                if JsonResponse.Get('value', JsonToken) then begin
                    JsonArray := JsonToken.AsArray();

                    foreach JsonTokenLoop in JsonArray do begin
                        JsonObjectLoop := JsonTokenLoop.AsObject();
                        if JsonObjectLoop.Get('name', JsonTokenLoop) then begin
                            JsonValue := JsonTokenLoop.AsValue();
                            if FileName = '' then begin
                                FileName := JsonValue.AsText();
                            end else begin
                                FileName := FileName + '\' + JsonValue.AsText();
                            end;
                        end;
                    end;
                    Message(FileName);
                end;

            end else begin
                //Report errors!
                HttpResponseMessage.Content.ReadAs(ResponseText);
                Error('Failed to fetch files from SharePoint: %1 %2', HttpResponseMessage.HttpStatusCode(), ResponseText);
            end;
        end else
            Error('Failed to send HTTP request to SharePoint');
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
}
