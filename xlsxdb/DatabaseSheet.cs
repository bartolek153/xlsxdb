using System.Data;
using System.Data.SqlClient;
using System.Security.Cryptography.X509Certificates;

using Microsoft.Toolkit.Uwp.Notifications;

using OfficeOpenXml;

using Serilog;

namespace xlsxdb;

internal class DatabaseSheet
{
    public string TableName { get; set; } = null!;
    public string FilePath { get; set; }
    public bool HasDatabaseObject { get; set; }
    public int NumRows { get; set; }
    public DataTable Data { get; set; }

    public static DatabaseSheet Read(string path)
    {
        DatabaseSheet dbsh = new DatabaseSheet();

        dbsh.FilePath = path;
        dbsh.TableName = Path.GetFileNameWithoutExtension(path).ToUpper();

        return dbsh;
    }

    public DatabaseSheet ValidateExistence(SqlConnection connection)
    {
        string query = $"SELECT COUNT(*) FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = '{TableName}'";

        using (SqlCommand command = new SqlCommand(query, connection))
        {
            int tableCount = (int)command.ExecuteScalar();

            HasDatabaseObject = tableCount > 0;
        }

        return this;
    }

    public void GenerateTableSql()
    {
        throw new NotImplementedException();

        //    string createQuery = $"CREATE TABLE {tableName} (";

        //    // starts at index 1, and desconsider table head line
        //    for (int col = 2; col <= ws.Dimension.Columns; col++)
        //    {
        //        //var cell = ws.Cells[2, col];

        //        //Console.WriteLine(cell.Value.GetType().ToString());
        //        //createQuery += "";
        //    }

        //    for (int row = 2; row <= ws.Dimension.Rows; row++)
        //    {
        //        var cell = ws.Cells[row, 4];

        //        Console.Write(cell.ToString() + "   ");
        //    }

        //    log.Warning("Comando a ser executado: ");

        //    Console.Write("Deseja prosseguir com a criação? (Y/N) ");
        //    var key = Console.ReadKey(true);
        //    Console.WriteLine(key.Key);
    }

    private void RetrieveTableStructure(SqlConnection conn)
    {
        this.Data = new DataTable();

        using (SqlDataAdapter dataAdapter = new SqlDataAdapter($"SELECT TOP 5 * FROM {TableName}", conn))
        {
            dataAdapter.Fill(this.Data);
            this.Data.Clear();
        }
    }

    private void RearrangeColumnOrder(ExcelWorksheet ws, DataTable dto)
    {
        var finalDt = new DataTable();

        for (int col = 1; col <= ws.Dimension.Columns; col++)
        {
            if (this.Data.Columns.Contains(ws.Cells[1, col].Text))
            {
                var sourceCol = this.Data.Columns[ws.Cells[1, col].Text];
                finalDt.Columns.Add(sourceCol!.ColumnName, sourceCol.DataType);
            }
        }

        this.Data = finalDt;
    }

    public DatabaseSheet Fill(SqlConnection conn)
    {
        RetrieveTableStructure(conn);

        using var package = new ExcelPackage(FilePath);
        var ws = package.Workbook.Worksheets[0];

        RearrangeColumnOrder(ws, this.Data);

        string[] datetimeFormats = { "yyyy-MM-dd", "MM/dd/yyyy", "dd/MM/yyyy" };

        for (int r = 2; r <= ws.Dimension.Rows; r++)
        {
            DataRow row = Data.NewRow();

            for (int c = 1; c <= ws.Dimension.Columns; c++)
            { 
                if (datetimeFormats.Any(dtf => ws.Cells[r, c].Style.Numberformat.Format.Contains(dtf)))
                {
                    row[c - 1] = ws.Cells[r, c].Text;
                    continue;
                }

                row[c - 1] = ws.Cells[r, c].Value;
            }

            Data.Rows.Add(row);
        }

        return this;
    }

    public DatabaseSheet CopyToDatabase(SqlConnection conn)
    {
        using SqlBulkCopy bulk = new SqlBulkCopy(conn);
        bulk.DestinationTableName = TableName;
        bulk.BatchSize = 10000;
        bulk.BulkCopyTimeout = 0;

        bulk.WriteToServer(this.Data);

        int rowsAffected = CountTableRows(this.TableName, conn);

        Notify("Processamento finalizado!", $"Tabela {TableName} carregada com {rowsAffected} registros.");

        return this;
    }

    public void DeleteFile()
    {
        File.Delete(FilePath);
    }

    private int CountTableRows(string tableName, SqlConnection conn)
    {
        string query = $"SELECT COUNT(*) FROM {tableName}";

        using SqlCommand command = new SqlCommand(query, conn);
        return (int)command.ExecuteScalar();
    }

    public static string[] GetExcelFiles()
    {
        return Directory.GetFiles(Directory.GetCurrentDirectory(), "*.xlsx");
    }

    public static (SqlConnection?, string) GetConnection(string connectionString)
    {
        try
        {
            var conn = new SqlConnection(connectionString);
            conn.Open();
            return (conn, "");
        }
        catch (Exception ex)
        {
            return (null, $"Não foi possível estabelecer uma conexão com a base: {ex.Message}");
        }
    }

    private void Notify(string title, string message)
    {
        new ToastContentBuilder()
                .AddText(title)
                .AddText(message)
                .Show();
    }

}
