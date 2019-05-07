using Syncfusion. XlsIO;
using System;
using System. IO;
using System.Linq;
namespace ImportFromCSV
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

                    //Preserve data types as per the value
                    application.PreserveCSVDataTypes = true;

                    //Read the CSV file
                    Stream csvStream = File.OpenRead(Path.GetFullPath(@"../../../TemplateSales.csv")); ;

                    //Reads CSV stream as a workbook
                    IWorkbook workbook = application.Workbooks.Open(csvStream);
                    IWorksheet sheet = workbook.Worksheets[0];

                    //Formatting the CSV data as a Table 
                    IListObject table = sheet.ListObjects.Create("SalesTable", sheet.UsedRange);
                    table.BuiltInTableStyle =  TableBuiltInStyles.TableStyleMedium6;
                    IRange location = table.Location;
                    location.AutofitColumns();

                    //Apply the proper latitude & longitude numerformat in the table
                    TryAndUpdateGeoLocation(table,"Latitude");
                    TryAndUpdateGeoLocation(table,"Longitude");

                    //Apply currency numberformat in the table column 'Price'
                    IRange columnRange = GetListObjectColumnRange(table,"Price");
                    if(columnRange != null)
                        columnRange.CellStyle.NumberFormat = "$#,##0.00";

                    //Apply Date time numberformat in the table column 'Transaction_date'
                    columnRange = GetListObjectColumnRange(table,"Transaction_date");
                    if(columnRange != null)
                        columnRange.CellStyle.NumberFormat = "m/d/yy h:mm AM/PM;@";

                    //Sort the data based on 'Products'
                    IDataSort sorter = table.AutoFilters.DataSorter;
                    ISortField sortField = sorter. SortFields. Add(0, SortOn. Values, OrderBy. Ascending);
                    sorter. Sort();

                    //Save the file in the given path
                    Stream excelStream;
                    excelStream = File.Create(Path.GetFullPath(@"../../../Output.xlsx"));
                    workbook.SaveAs(excelStream);
                    excelStream.Dispose();
                }
            }
            catch(Exception e)
            {
                Console.WriteLine("Unexpected Exception:"+e.Message);
            }
        }

        private static void TryAndUpdateGeoLocation(IListObject table, string unitString)
        {
            IRange columnRange = GetListObjectColumnRange(table, unitString);
            if(columnRange == null) return;
            columnRange.Worksheet.EnableSheetCalculations();
            foreach(IRange range in columnRange.Cells)
            {
                string currentValue = range.Value;
                range.Value2 = "=TEXT(TRUNC("+currentValue+"), \"0\" & CHAR(176) & \" \") &" +
                    " TEXT(INT((ABS("+currentValue+")- INT(ABS("+currentValue+")))*60), \"0' \") " +
                    "& TEXT(((((ABS("+currentValue+")-INT(ABS("+currentValue+")))*60)-" +
                    " INT((ABS("+currentValue+") - INT(ABS("+currentValue+")))*60))*60), \" 0''\")";
            }
        }

        private static IRange GetListObjectColumnRange(IListObject table, string name)
        {
            IListObjectColumn column = table.Columns.FirstOrDefault(x=> x.Name.Contains(name, StringComparison.InvariantCultureIgnoreCase));
            if(column!=null)
            {
                IRange location = table.Location;
                return location.Worksheet[location.Row+1, location.Column + column.Index - 1, location.LastRow, location.Column + column.Index - 1 ];
            }
            else
                return null;
        }
    }
}
