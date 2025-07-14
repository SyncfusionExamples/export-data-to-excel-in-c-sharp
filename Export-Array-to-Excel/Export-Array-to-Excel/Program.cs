using Syncfusion. XlsIO;
using System;
using System.IO;

namespace ImportFromArray
{
    class Program
    {
        static void Main(string[] args)
        {

            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;

                //Reads input Excel stream as a workbook
                IWorkbook workbook = application.Workbooks.Open(File.OpenRead(Path.GetFullPath(@"../../../Expenses.xlsx")));
                IWorksheet sheet = workbook.Worksheets[0];

                //Preparing first array with different data types
                object[] expenseArray = new object[14]
                {"Paul Pogba", 469.00d, 263.00d, 131.00d, 139.00d, 474.00d, 253.00d, 467.00d, 142.00d, 417.00d, 324.00d, 328.00d, 497.00d, "=SUM(B11:M11)"};

                //Inserting a new row by formatting as a previous row.
                sheet.InsertRow(11, 1, ExcelInsertOptions.FormatAsBefore);

                //Import Peter's expenses and fill it horizontally
                sheet.ImportArray(expenseArray, 11, 1, false);

                //Preparing second array with double data type
                double[] expensesOnDec = new double[6]
                {179.00d, 298.00d, 484.00d, 145.00d, 20.00d, 497.00d};

                //Modify the December month's expenses and import it vertically
                sheet.ImportArray(expensesOnDec, 6, 13, true);

                //Save the file in the given path
                Stream excelStream = File.Create(Path.GetFullPath(@"Output.xlsx"));
                workbook.SaveAs(excelStream);
                excelStream.Dispose();
            }

        }
    }
}
