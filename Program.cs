using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;

string[] scopes = { SheetsService.Scope.Spreadsheets };
const string applicationName = "App-Name";
const string spreadsheetId = "1_ecHRnXzm-ZgtRRzbr99EsWBFTtd-13zj73dvKJJ1lY";
const string sheet = "sample-sheet";
GoogleCredential credential;

using (var stream = new FileStream("client_secrets.json", FileMode.Open, FileAccess.Read))
{
    credential = GoogleCredential.FromStream(stream).CreateScoped(scopes);
}

var service = new SheetsService(new BaseClientService.Initializer()
{
    HttpClientInitializer = credential,
    ApplicationName = applicationName,
});

createEntry();
readEntries();
DeleteEntry();
updateEntry();

void readEntries()
{
    var range = $"{sheet}!A1:C6";
    var request = service.Spreadsheets.Values.Get(spreadsheetId, range);
    var response = request.Execute();
    var values = response.Values;
    if (values is { Count: > 0 })
    {
        foreach (var row in values)
        {
            Console.WriteLine("{0} | {1}", row[1], row[2]);
        }
    }
    else
    {
        Console.WriteLine("No Data Found!");
    }
}

void createEntry()
{
    var range = $"{sheet}!A:F";
    var valueRange = new ValueRange();
    
    var objectList = new List<object>() { "Hello", "This", "is", "a", "test!!" };
    valueRange.Values = new List<IList<object>> { objectList };

    var appendRequest = service.Spreadsheets.Values.Append(valueRange, spreadsheetId, range);
    appendRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.USERENTERED;
    var appendResponse = appendRequest.Execute();
}

void updateEntry()
{
    var range = $"{sheet}!A10:F10";
    var valueRange = new ValueRange();
    
    var objectList = new List<object> { "Hello", "This", "is", "updated", "entry!!", "new@@@@@" };
    valueRange.Values = new List<IList<object>> { objectList };

    var updateRequest = service.Spreadsheets.Values.Update(valueRange, spreadsheetId, range);
    updateRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;
    var appendResponse = updateRequest.Execute();
}

void DeleteEntry()
{
    var range = $"{sheet}!A8:E8";
    var requestBody = new ClearValuesRequest();

    var deleteResponse = service.Spreadsheets.Values.Clear(requestBody, spreadsheetId, range);
    deleteResponse.Execute();
}