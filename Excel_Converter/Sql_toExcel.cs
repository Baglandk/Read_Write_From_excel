using System.Data;
using System.Data.SqlClient;
using ClosedXML.Excel;

internal class Sql_toExcel
{
    internal DataTable FetchSql(string connectionString, string query)
    {
       using(SqlConnection connection = new SqlConnection(connectionString))
        {
            using(SqlCommand command = new SqlCommand(query, connection))
            {
                connection.Open();
                DataTable dataTable = new DataTable();
                using(SqlDataReader reader = command.ExecuteReader())
                {
                    dataTable.Load(reader);
                }
                return dataTable;
            }
        }
    }

    internal void Export_toExcel(DataTable tble2, string filePath)
    {
        using(XLWorkbook workBook = new XLWorkbook())
        {
            workBook.AddWorksheet(tble2, "Table1 ");
            workBook.SaveAs(filePath);
        }
    }
}