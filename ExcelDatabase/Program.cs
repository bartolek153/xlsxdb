
using OfficeOpenXml;
using Serilog;
using System.Data;
using System.Data.SqlClient;


namespace ExcelDatabase;

class Program
{
    static void Main(string[] args)
    {
        // paths
        string tableName;
        string path = Directory.GetCurrentDirectory();
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

        // connection settings
        string connectionString = "Server=(localdb)\\MSSQLLocalDB;Initial Catalog=augustagourmetdb2;Integrated Security=true";
        using SqlConnection conn = new SqlConnection(connectionString);
        conn.Open();


        var log = new LoggerConfiguration()
            .WriteTo.Console()
            //.WriteTo.File("log.txt")
            .CreateLogger();

        //log.Information("Starting at " + DateTime.Now.ToString());
        
        Console.WriteLine();

        // get Excel file paths
        var files = Directory.GetFiles(path, "*.xlsx");

        // iter each file (from absolute path)
        foreach (var filePath in files)
        {
            tableName = Path.GetFileNameWithoutExtension(filePath);
            using var package = new ExcelPackage(filePath);
            var ws = package.Workbook.Worksheets[0];

            // validate the table actually exists
            try
            {
                //tableName += 'a';

                using (SqlCommand cmd = new SqlCommand($"SELECT TOP 1 * FROM {tableName}", conn))
                    cmd.ExecuteScalar();
            }
            catch
            {
                log.Warning($"Tabela *{tableName}* não existe no banco.");

                string createQuery = $"CREATE TABLE {tableName} (";

                // starts at index 1, and desconsider table head line
                for (int col = 2; col <= ws.Dimension.Columns; col++)
                {
                    //var cell = ws.Cells[2, col];

                    //Console.WriteLine(cell.Value.GetType().ToString());
                    //createQuery += "";
                }

                for (int row = 2; row <= ws.Dimension.Rows; row++)
                {
                    var cell = ws.Cells[row, 4];

                    Console.Write(cell.ToString() + "   ");   
                }

                log.Warning("Comando a ser executado: ");

                Console.Write("Deseja prosseguir com a criação? (Y/N) ");
                var key = Console.ReadKey(true);
                Console.WriteLine(key.Key);
            }

            log.Information($"Tabela {tableName} criada na database.");
            log.Information("Loading data...");

            DataTable data = new DataTable();

            try
            {
                using (SqlDataAdapter dataAdapter = new SqlDataAdapter($"SELECT TOP 10 * FROM {tableName}", conn))
                {
                    dataAdapter.Fill(data);
                    data.Clear();
                }

                // check if there is a company column
                var head = ws.Cells["1:1"].First(c => c.Value.ToString() == "companhia").Start.Column;

                if (head == 0)
                {
                    log.Fatal($"Colua companhia em {tableName} está faltando no xlsx.");
                    break;
                }
                else if (head != 0 && head != 1)
                {
                    log.Fatal("Colunas precisam estar na mesma ordem");
                    break;
                }

                for (int r = 2; r <= ws.Dimension.Rows; r++)
                {
                    DataRow row = data.NewRow();

                    for (int c = 1; c <= ws.Dimension.Columns; c++)
                    {
                        object value = ws.Cells[r, c].Value;

                        if (value is DateTime dateTimeValue)
                        {
                            row[c - 1] = dateTimeValue;
                            continue;
                        }

                        row[c - 1] = value;
                    }

                    data.Rows.Add(row);
                }


            }
            catch (Exception ex)
            {

                throw;
            }

            using (SqlBulkCopy bulk = new SqlBulkCopy(conn))
            {
                bulk.DestinationTableName = tableName;
                bulk.BatchSize = 10000;
                bulk.BulkCopyTimeout = 0;

                bulk.WriteToServer(data);
            }
        }
    }
}
