using Syncfusion. XlsIO;
using System;
using System. Collections. Generic;
using System. Diagnostics;
using System. Drawing;
using System. IO;
using System. Linq;
using System. Reflection;
using System. Text;
using System. Threading. Tasks;
using System. Xml. Linq;

namespace ImportFromCollectionObjects
{
    class Program
    {
        static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Excel2016;

                //Read the data from XML file
                StreamReader reader = new StreamReader(Path.GetFullPath(@"../../Data/Customers.xml"));

                //Assign the data to the customerObjects collection
                IEnumerable<Customer> customerObjects = GetData <Customer>(reader.ReadToEnd());   

                //Create a new workbook
                IWorkbook workbook = application.Workbooks.Create(1);
                IWorksheet sheet = workbook.Worksheets[0];

                //Import data from customerObjects collection
                sheet.ImportData(customerObjects, 5, 1, false);

                #region Define Styles
                IStyle pageHeader = workbook.Styles.Add("PageHeaderStyle");
                IStyle tableHeader = workbook.Styles.Add("TableHeaderStyle");

                pageHeader.Font.RGBColor = Color.FromArgb(0, 83, 141, 213);
                pageHeader.Font.FontName = "Calibri";
                pageHeader.Font.Size = 18;
                pageHeader.Font.Bold = true;
                pageHeader.HorizontalAlignment = ExcelHAlign.HAlignCenter;
                pageHeader.VerticalAlignment = ExcelVAlign.VAlignCenter;

                tableHeader.Font.Color = ExcelKnownColors.White;
                tableHeader.Font.Bold = true;
                tableHeader.Font.Size = 11;
                tableHeader.Font.FontName = "Calibri";
                tableHeader.HorizontalAlignment = ExcelHAlign.HAlignCenter;
                tableHeader.VerticalAlignment = ExcelVAlign.VAlignCenter;
                tableHeader.Color = Color.FromArgb(0, 118, 147, 60);
                tableHeader.Borders[ExcelBordersIndex.EdgeLeft].LineStyle = ExcelLineStyle.Thin;
                tableHeader.Borders[ExcelBordersIndex.EdgeRight].LineStyle = ExcelLineStyle.Thin;
                tableHeader.Borders[ExcelBordersIndex.EdgeTop].LineStyle = ExcelLineStyle.Thin;
                tableHeader.Borders[ExcelBordersIndex.EdgeBottom].LineStyle = ExcelLineStyle.Thin;
                #endregion

                #region Apply Styles
                //Apply style to the header
                sheet["A1"].Text = "Yearly Sales Report";
                sheet["A1"].CellStyle = pageHeader;

                sheet["A2"].Text = "Namewise Sales Comparison Report";
                sheet["A2"].CellStyle = pageHeader;
                sheet["A2"].CellStyle.Font.Bold = false;
                sheet["A2"].CellStyle.Font.Size = 16;

                sheet["A1:D1"].Merge();
                sheet["A2:D2"].Merge();
                sheet["A3:A4"].Merge();
                sheet["D3:D4"].Merge();
                sheet["B3:C3"].Merge();

                sheet["B3"].Text = "Sales";
                sheet["A3"].Text = "Sales Person";
                sheet["B4"].Text = "January - June";
                sheet["C4"].Text = "July - December";
                sheet["D3"].Text = "Change(%)";
                sheet["A3:D4"].CellStyle = tableHeader;
                sheet.UsedRange.AutofitColumns();
                sheet.Columns[0].ColumnWidth = 24;
                sheet.Columns[1].ColumnWidth = 21;
                sheet.Columns[2].ColumnWidth = 21;
                sheet.Columns[3].ColumnWidth = 16;
                #endregion

                sheet.UsedRange.AutofitColumns();

                //Save the file in the given path
                Stream excelStream = File.Create(Path.GetFullPath(@"Output.xlsx"));
                workbook.SaveAs(excelStream);
                excelStream.Dispose();
                Process.Start(Path.GetFullPath(@"Output.xlsx"));
            }
        }

        internal static IEnumerable<T> GetData<T>(string xml)
        where T : Customer, new()
        {
            return XElement.Parse(xml)
               .Elements("Customers")
               .Select(c => new T
               {
                    SalesPerson = (string)c.Element("SalesPerson"),
                    SalesJanJune = int.Parse((string)c.Element("SalesJanJune")),
                    SalesJulyDec=  int.Parse((string)c.Element("SalesJulyDec")),
                    Change=  int.Parse((string)c.Element("Change"))
               });
        }
   
    }
}
