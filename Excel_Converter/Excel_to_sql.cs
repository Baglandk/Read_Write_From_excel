using ExcelDataReader;
using Microsoft.Data.SqlClient;
using System.Data;

internal class Excel_to_sql
{
    internal DataTable ImportDataFromExcel(string excelPath)
    {
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
        using(var stream = File.Open(excelPath, FileMode.Open, FileAccess.Read))
        {
            using(var reader = ExcelReaderFactory.CreateReader(stream))
            {
                var result = reader.AsDataSet(new ExcelDataSetConfiguration
                {
                    ConfigureDataTable = _ => new ExcelDataTableConfiguration
                    {
                        UseHeaderRow = true // Set this to false if your Excel file doesn't have a header row
                    }
                });
                return result.Tables[0];
            }
        }
    }
    internal void InsertIntoSql(DataTable dataTable, string connectionString)
    {
        using(SqlConnection connection = new SqlConnection(connectionString))
        {
            connection.Open();
            using(SqlBulkCopy bulkCopy = new SqlBulkCopy(connection))
            {
                bulkCopy.DestinationTableName = "Table1";
                bulkCopy.WriteToServer(dataTable);
            }
        }
    }
}