using System;
using System.Data;

class Program
{
    static void Main()
    {
        //string excelPath = "/Users/baglan/Desktop/excel_tosql.xlsx";
        string connectionString = @"Server=localhost; User Id=SA; Password=reallyStrongPwd123;
                                       Initial Catalog=PhoneBook;TrustServerCertificate=True;
                                        Integrated Security=False";

        string filePath = "/Users/baglan/Desktop/Table1.xlsx";
        string query = "SELECT * FROM Table1";
        //Excel_to_sql s1 = new Excel_to_sql();
        //DataTable tbale = s1.ImportDataFromExcel(excelPath);
        //s1.InsertIntoSql(tbale, connectionString);

        Sql_toExcel q1 = new Sql_toExcel();
        DataTable tble2 = q1.FetchSql(connectionString, query);
        q1.Export_toExcel(tble2, filePath);
    }
}