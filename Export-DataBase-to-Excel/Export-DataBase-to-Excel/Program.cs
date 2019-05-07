using Syncfusion. XlsIO;
using System;
using System. Collections. Generic;
using System. Data. OleDb;
using System. Diagnostics;
using System. IO;
using System. Linq;
using System. Text;
using System. Threading. Tasks;

namespace ImportFromDataBase
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                using (ExcelEngine excelEngine = new ExcelEngine())
                {
                    IApplication application = excelEngine.Excel;
                    application.DefaultVersion = ExcelVersion.Excel2016;

                    //Create a new workbook
                    IWorkbook workbook = application.Workbooks.Create(1);
                    IWorksheet sheet = workbook.Worksheets[0];
                    
                    if(sheet.ListObjects.Count == 0)
                    {
                        //Estabilishing the connection in the worksheet
                        string dBPath = Path.GetFullPath(@"../../Data/EmployeeData.mdb");
                        string ConnectionString = "OLEDB;Provider=Microsoft.JET.OLEDB.4.0;Password=\"\";User ID=Admin;Data Source="+ dBPath;
                        string query = "SELECT EmployeeID,FirstName,LastName,Title,HireDate,Extension,ReportsTo FROM [Employees]";
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
                    Process.Start(Path.GetFullPath(@"Output.xlsx"));
                }
            }
            catch(Exception e)
            {
                Console.WriteLine("Unexpected Exception:"+e.Message);
            }
        }
    }
}
