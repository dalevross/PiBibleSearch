// See https://aka.ms/new-console-template for more information


using System.Data;
using System.Data.OleDb;
using System.Diagnostics.Eventing.Reader;
using System.Net.Http;
using System.Reflection.Metadata;
using System.Text.Json;
using System.Threading.Tasks;

string excelConnectString = @"Provider=Microsoft.Jet.OLEDB.4.0; Data Source=bibletaxonomy.xls;Extended Properties=""Excel 8.0;HDR=YES;""";
//string excelConnectString = @"Provider = Microsoft.Jet.OLEDB.4.0;Data Source = " + excelFileName + ";" + "Extended Properties = Excel 8.0; HDR=Yes;IMEX=1";

OleDbConnection objConn = new OleDbConnection(excelConnectString);

DataTable Contents = new DataTable();
using (OleDbDataAdapter adapter = new OleDbDataAdapter("Select * From [Sheet1$]", objConn))
{
    adapter.Fill(Contents);
}
//Console.WriteLine(Contents.Rows[0][0]);

var apiClient = new ApiClient();
var jsonParser = new JsonParser();

using(StreamWriter sw = new StreamWriter("verses1832.txt"))
{

    foreach(DataRow row in Contents.Rows)
    {
        

        string book = row[0].ToString()
        .Replace("1","1st")
        .Replace("2","2nd")
        .Replace("3","3rd");
        if(book == null)
            break;

        int chapter = Int32.Parse(row[1].ToString());

        int verse = Int32.Parse(row[2].ToString());

        string verseString = $"{book} {chapter}: {verse}";
        Console.WriteLine(verseString);
        string digits = "1832";
        //if(verseString.Contains("3") && verseString.Contains("1") && verseString.Contains("4"))
        if(digits.All(letter => $"{chapter} {verse}".Contains(letter)))
        {
            string url = $"http://localhost:3000/api/v1/verse?book={book}&chapter={chapter}&verses={verse}&version=KJV";

            
            try{
                string jsonResponse = await apiClient.GetApiResponseAsync(url);
                ApiResponse apiResponse = jsonParser.ParseJson(jsonResponse);

                if (apiResponse != null)
                {
                    // Use the deserialized object
                    sw.WriteLine($"{verseString}:{apiResponse.passage}");
                } 
            }
            catch
            {
                Console.WriteLine(url);
            }
        }
    }


}
Console.WriteLine("Done");
//http://localhost:3000/api/v1/verse?book=John&chapter=3&verses=16&version=NLT

    



public class ApiClient
{
    private readonly HttpClient _httpClient;

    public ApiClient()
    {
        _httpClient = new HttpClient();
    }

    public async Task<string> GetApiResponseAsync(string url)
    {
        HttpResponseMessage response = await _httpClient.GetAsync(url);
        response.EnsureSuccessStatusCode();
        return await response.Content.ReadAsStringAsync();
    }
}

public class ApiResponse
{
    // Define properties according to the JSON structure
    public string citation { get; set; }
    public string passage { get; set; }
    // ... other properties
}

public class JsonParser
{
    public ApiResponse ParseJson(string jsonString)
    {
        try
        {
            return JsonSerializer.Deserialize<ApiResponse>(jsonString);
        }
        catch (JsonException e)
        {
            Console.WriteLine("JSON Parsing Error: " + e.Message);
            return null;
        }
    }
}



