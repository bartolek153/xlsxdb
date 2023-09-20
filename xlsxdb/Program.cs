using dotenv.net;

using Microsoft.Toolkit.Uwp.Notifications;

using OfficeOpenXml;

using Serilog;
using Serilog.Core;

using Spectre.Console;
using Spectre.Console.Extensions;

namespace xlsxdb;

class Program
{
    static async Task Main(string[] args)
    {
        AnsiConsole.Write(
            new FigletText("xlsxdb")
                .Centered()
                .Color(Color.Orange1));

        Logger log = new LoggerConfiguration()
                   .WriteTo.Console()
                   .CreateLogger();

        await AnsiConsole.Status()
            .Spinner(Spinner.Known.Star)
            .SpinnerStyle(Style.Parse("green"))
            .Start("xlsxdb init", async ctx =>
            {
                AnsiConsole.MarkupLine("[invert green]Hello World[/]");
                Console.WriteLine("call1");

                var a = new Test();
                await a.TestMethod(); Console.WriteLine("end call1");

                // log.Information("Initializing ...");

                DotEnv.Load();
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                ctx.Spinner(Spinner.Known.Dots11);
                ctx.Status("Conectando-se ao banco...");

                (var conn, string error) = DatabaseSheet.GetConnection(Environment.GetEnvironmentVariable("CONNECTION_STR")!);
                if (conn is null)
                {
                    log.Fatal(error);
                    Console.ReadKey();

                    return;
                }

                ctx.Status("Conectado!");
                Thread.Sleep(1000);


                ctx.Status("Lendo arquivos...");
                string[] filePaths = DatabaseSheet.GetExcelFiles();

                if (filePaths.Length == 0)
                {
                    new ToastContentBuilder()
                        .AddText("Processamento abortado!")
                        .AddText($"Não há arquivos (.xlsx) a serem carregados, no diretório atual.")
                        .Show();

                    Console.ReadKey();
                    return;
                }

                // iter each file (abs path)
                foreach (var path in filePaths)
                {
                    try
                    {
                        var ws = DatabaseSheet
                        .Read(path)
                        .ValidateExistence(conn);

                        if (!ws.HasDatabaseObject)
                        {
                            //log.Warning($"Tabela *{ws.TableName}* não existe na base.");
                            ws.GenerateTableSql();
                        }

                        ctx.Status("Carregando dados...");

                        ws
                            .Fill(conn)
                            .CopyToDatabase(conn)
                            .DeleteFile();
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine();
                        ctx.Status("Erro: " + ex.Message);

                        new ToastContentBuilder()
                        .AddText("Erro!")
                        .AddText($"Veja a saída no console.")
                        .Show();

                        Console.ReadKey();
                    }
                }
            });

        log.Information("Processo finalizado ...");
        Thread.Sleep(1000);
        Console.ReadKey();
    }

    public void Log()
    {
        AnsiConsole.Markup("[green]This is all green[/]");
    }
}
