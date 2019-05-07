using System;
using System. Collections. Generic;
using System. Linq;
using System. Text;
using System. Threading. Tasks;
using Syncfusion. XlsIO;
using System. Diagnostics;
using System. IO;
using System.Data;
namespace ImportFromDataTable
{
    class Program
    {
        static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Excel2016;

                //Create a new workbook
                IWorkbook workbook = application.Workbooks.Create(1);
                IWorksheet sheet = workbook.Worksheets[0];

                //Create a dataset from XML file
                DataSet customersDataSet = new DataSet();
                customersDataSet.ReadXml(Path.GetFullPath(@"../../Data/Employees.xml"));

                //Create datatable from the dataset
                DataTable dataTable = new DataTable();
                dataTable = customersDataSet.Tables[1];

                //Import data from the data table with column header, at first row and first column, and by its column type.
                sheet.ImportDataTable(dataTable, true, 1, 1, true);

                //Creating Excel table or list object and apply style to the table
                IListObject table = sheet.ListObjects.Create("Employee_PersonalDetails", sheet.UsedRange);

                table.BuiltInTableStyle = TableBuiltInStyles.TableStyleMedium14;

                //Autofit the columns
                sheet.UsedRange.AutofitColumns();

                //Save the file in the given path
                Stream excelStream = File.Create(Path.GetFullPath(@"Output.xlsx"));
                workbook.SaveAs(excelStream);
                excelStream.Dispose();
                Process.Start(Path.GetFullPath(@"Output.xlsx"));
            }
        }
    }
}
