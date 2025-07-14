using Syncfusion.XlsIO;
using System;
using System. IO;
using System.Linq;
namespace ImportFromCSV
{
    class Program
    {
        static void Main(string[] args)
        {

            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;

                //Preserve data types as per the value
                application.PreserveCSVDataTypes = true;

                //Read the CSV file
                Stream csvStream = File.OpenRead(Path.GetFullPath(@"../../../TemplateSales.csv")); ;

                //Reads CSV stream as a workbook
                IWorkbook workbook = application.Workbooks.Open(csvStream);
                IWorksheet sheet = workbook.Worksheets[0];

                //Formatting the CSV data as a Table 
                IListObject table = sheet.ListObjects.Create("SalesTable", sheet.UsedRange);
                table.BuiltInTableStyle = TableBuiltInStyles.TableStyleMedium6;
                IRange location = table.Location;
                location.AutofitColumns();

                //Sort the data based on 'Products'
                IDataSort sorter = table.AutoFilters.DataSorter;
                ISortField sortField = sorter.SortFields.Add(0, SortOn.Values, OrderBy.Ascending);
                sorter.Sort();

                //Save the file in the given path
                Stream excelStream;
                excelStream = File.Create(Path.GetFullPath(@"../../../Output.xlsx"));
                workbook.SaveAs(excelStream);
                excelStream.Dispose();
            }

        }
    }
}
