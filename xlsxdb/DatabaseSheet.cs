using System.Data;
using System.Data.SqlClient;
using System.Security.Cryptography.X509Certificates;

using Microsoft.Toolkit.Uwp.Notifications;

using OfficeOpenXml;

using Serilog;

namespace xlsxdb;

internal class DatabaseSheet
{
    //
    // TODO: Add other drivers support
    //

    // the sql connection object (only supports SQL Server)
    public static SqlConnection? Connection { get; set; }

    public string TableName { get; set; } = null!;

    public string FilePath { get; set; } = null!;

    // indicates if the destiny table
    // exists or not in the database
    public bool HasDatabaseObject { get; set; }

    // spreadsheet column name
    public string[] Header { get; set; } = null!;

    public int NumRows { get; set; }

    public DataTable Data { get; set; } = null!;

    public DatabaseSheet() { }
    public DatabaseSheet(string path, string tableName) 
    {
        TableName = tableName;
        FilePath = path;
    }

    // Establishes connection with the database, ping this connection
    // and set the result object in Connection property,
    // that can be used throughout the execution.
    public static (SqlConnection?, string) SetConnection(string connectionString)
    {
        try
        {
            var conn = new SqlConnection(connectionString);
            conn.Open();

            Connection = conn;

            return (conn, "");
        }
        catch (Exception ex)
        {
            return (null, $"Could not establish connection with the database. Reason: {ex.Message}");
        }
    }

    // Returns an instantiated DatabaseSheet object,
    // with metadata loaded as properties.
    public static DatabaseSheet Read(string path)
    {
        return new DatabaseSheet(
            path, 
            Path.GetFileNameWithoutExtension(path).ToUpper()
        );
    }

    public DatabaseSheet ValidateExistence()
    {
        string query = $"SELECT COUNT(*) FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = '{TableName}'";

        using (SqlCommand command = new SqlCommand(query, Connection))
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

    // Retrieve a DataTable object 
    // that contains the column names of destiny table
    // in the database
    private void RetrieveTableStructure(SqlConnection conn)
    {
        this.Data = new DataTable();

        using (SqlDataAdapter dataAdapter = new SqlDataAdapter($"SELECT TOP 5 * FROM {TableName}", conn))
        {
            dataAdapter.Fill(this.Data);
            this.Data.Clear();
        }
    }

    private void RearrangeColumnOrder(DataTable dto)
    {
        var ordinalMap = dto.Columns.Cast<DataColumn>()
            .Select((col, index) => new { Column = col, Ordinal = index })
            .ToDictionary(
                item => item.Column.ColumnName,
                item => item.Ordinal);

        var newColumns = Header
            .Select(columnName => ordinalMap[columnName])
            .ToList();

        foreach (var newOrdinal in newColumns) 
        {
            dto.Columns[newOrdinal].SetOrdinal(newColumns.IndexOf(newOrdinal));
        }
    }

    public DatabaseSheet Fill()
    {
        RetrieveTableStructure(Connection!);

        using var package = new ExcelPackage(FilePath);
        var ws = package.Workbook.Worksheets[0];

        foreach (var headerCell in ws.Cells[1, 1, 1, ws.Dimension.Columns])
        {
            if (headerCell.Value is not null)
                Header.Append(headerCell.Value.ToString());
        }

        RearrangeColumnOrder(this.Data);

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

    public DatabaseSheet CopyToDatabase()
    {
        using SqlBulkCopy bulk = new SqlBulkCopy(Connection!);
        bulk.DestinationTableName = TableName;
        bulk.BatchSize = 10000;
        bulk.BulkCopyTimeout = 0;

        bulk.WriteToServer(this.Data);

        int rowsAffected = CountTableRows(this.TableName, Connection!);

        Utils.Notify("Processamento finalizado!", $"Tabela {TableName} carregada com {rowsAffected} registros.");

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
}
