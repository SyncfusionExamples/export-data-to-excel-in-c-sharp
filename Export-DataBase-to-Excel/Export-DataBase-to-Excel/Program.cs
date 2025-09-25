using Syncfusion.XlsIO;
using System.IO;

namespace ImportFromDataBase
{
    class Program
    {
        static void Main(string[] args)
        {

            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;

                //Create a new workbook
                IWorkbook workbook = application.Workbooks.Create(1);
                IWorksheet sheet = workbook.Worksheets[0];

                if (sheet.ListObjects.Count == 0)
                {
                    //Establishing the connection in the worksheet
                    string dBPath = Path.GetFullPath(@"../../Data/EmployeeData.mdb");
                    string ConnectionString = "OLEDB;Provider=Microsoft.JET.OLEDB.4.0;Password=\"\";User ID=Admin;Data Source=" + dBPath;
                    string query = "SELECT EmployeeID,FirstName,LastName,Title,HireDate,City,Extension FROM [Employees]";
                    IConnection Connection = workbook.Connections.Add("Connection1", "Sample connection with MsAccess", ConnectionString, query, ExcelCommandType.Sql);
                    sheet.ListObjects.AddEx(ExcelListObjectSourceType.SrcQuery, Connection, sheet.Range["A1"]);
                }

                //Refresh Excel table to get updated values from database
                sheet.ListObjects[0].Refresh();

                sheet.UsedRange.AutofitColumns();

                //Save the file in the given path
                Stream excelStream = File.Create(Path.GetFullPath(@"Output.xlsx"));
                workbook.SaveAs(excelStream);
                excelStream.Dispose();
            }

        }

    }
}
