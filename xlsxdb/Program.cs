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
    static void Main(string[] args)
    {
        AnsiConsole.Write(
            new FigletText("xlsxdb")
                .Centered()
                .Color(Color.Orange1));

        AnsiConsole.Status()
            .Spinner(Spinner.Known.Star)
            .SpinnerStyle(Style.Parse("green"))
            .Start("xlsxdb init", ctx =>
            {
                AnsiConsole.MarkupLine("[invert green]Hello World[/]");
                Console.WriteLine("call1");

                //log.Information("Initializing ...");

                DotEnv.Load();
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                ctx.Spinner(Spinner.Known.Dots11);
                ctx.Status("Conectando-se ao banco...");

                (var conn, string error) = DatabaseSheet.SetConnection(Environment.GetEnvironmentVariable("CONNECTION_STR")!);
                if (conn is null)
                {
                    //log.Fatal(error);
                    Console.ReadKey();

                    return;
                }

                ctx.Status("Conectado!");
                Thread.Sleep(1000);


                ctx.Status("Lendo arquivos...");
                string[] filePaths = Utils.GetFilesByExtension("xlsx");

                if (filePaths.Length == 0)
                {
                    Utils.Notify("Processamento abortado!", $"Não há arquivos (.xlsx) a serem carregados, no diretório atual.");
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
                            .ValidateExistence();

                        if (!ws.HasDatabaseObject)
                        {
                            //log.Warning($"Tabela *{ws.TableName}* não existe na base.");
                            ws.GenerateTableSql();
                        }

                        ctx.Status("Carregando dados...");

                        ws
                            .Fill()
                            .CopyToDatabase()
                            .DeleteFile();
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine();
                        ctx.Status("Erro: " + ex.Message);
                        Utils.Notify("Erro!", $"Veja a saída no console.");
                        Console.ReadKey();
                    }
                }
            });

        //log.Information("Processo finalizado ...");
        Thread.Sleep(1000);
        Console.ReadKey();
    }
}
