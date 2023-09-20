
namespace xlsxdb;

public class Test
{
    public async Task<string> TestMethod()
    {
        Console.WriteLine("call2");

        var client = new HttpClient();

        // var response = await client.GetAsync("https://geek-jokes.sameerkumar.website/api?format=json");
        // var result = await response.Content.ReadAsStringAsync();

        Thread.Sleep(1000);


        Console.WriteLine("end call2");

        return "asdf";
    }
}