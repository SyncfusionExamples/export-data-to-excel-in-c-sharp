# Export Data to Excel in C#
<a href="https://www.syncfusion.com/excel-framework/net"><strong>Syncfusion Excel (XlsIO) library</strong></a> is a .NET Excel library that allows the user to export data to Excel in C# and VB.NET from various data sources like data tables, arrays, collections of objects, databases, CSV/TSV, and Microsoft Grid controls in a very simple and easy way. Exporting data to Excel helps in visualizing the data in a more understandable fashion. This feature helps to generate financial reports, banking statements, and invoices, while also allowing for filtering large data, validating data, formatting data, and more.

You can refer the <a href="https://help.syncfusion.com/file-formats/xlsio/working-with-data?_ga=2.120276040.1381167263.1557135100-214292665.1551328372#importing-data-to-worksheets">documention</a> to know more in detail.

Essential XlsIO provides the following ways to export data to Excel:

1. DataTable to Excel
2. Collection of objects to Excel
3. Database to Excel
4. Microsoft Grid controls to Excel
5. Array to Excel
6. CSV to Excel

This repository contains the examples of these methods and explains how to implement them.

<h2 id="DataTable-to-Excel">1. Export from DataTable to Excel</h2>

Data from <a href="https://en.wikipedia.org/wiki/ADO.NET">ADO.NET</a> objects such as <a href="https://docs.microsoft.com/en-us/dotnet/framework/data/adonet/dataset-datatable-dataview/datatables">datatable</a>, <a href="https://docs.microsoft.com/en-us/dotnet/api/system.data.datacolumn">datacolumn</a>, and <a href="https://docs.microsoft.com/en-us/dotnet/framework/data/adonet/dataset-datatable-dataview/dataviews">dataview</a> can be exported to Excel worksheets. The exporting can be done as column headers, by recognizing column types or cell value types, as hyperlinks, and as large dataset, all in a few seconds.

<strong>Exporting DataTable to Excel</strong> worksheets can be achieved through the <a href="https://help.syncfusion.com/file-formats/xlsio/working-with-data#import-data-from-datatable">ImportDataTable</a> method. The following code sample shows how to export a datatable of employee details to an Excel worksheet.

```csharp
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
```

<img class="aligncenter wp-image-38671 size-full" src="https://blog.syncfusion.com/wp-content/uploads/2019/04/DataTable-to-Excel_900x177.png" sizes="(max-width: 900px) 100vw, 900px" srcset="https://blog.syncfusion.com/wp-content/uploads/2019/04/DataTable-to-Excel_900x177.png 900w, https://blog.syncfusion.com/wp-content/uploads/2019/04/DataTable-to-Excel_900x177-300x59.png 300w, https://blog.syncfusion.com/wp-content/uploads/2019/04/DataTable-to-Excel_900x177-768x151.png 768w" alt="Export DataTable to Excel in C#" width="900" height="177" />

<em>Output of DataTable to Excel</em>


When exporting large data to Excel, and if there is no need to apply number formats and styles, you can make use of the <a href="https://help.syncfusion.com/cr/cref_files/windowsforms/Syncfusion.XlsIO.Base~Syncfusion.XlsIO.IWorksheet~ImportDataTable(DataTable,Int32,Int32,Boolean).html">ImportDataTable</a> overload with the TRUE value for <em>importOnSave</em> parameter. Here, the export happens while saving the Excel file.

Use this option to export large data with high performance.

```csharp
value = instance.ImportDataTable(dataTable, firstRow, firstColumn, importOnSave)
```

If you have a named range and like to <a href="https://help.syncfusion.com/cr/cref_files/windowsforms/Syncfusion.XlsIO.Base~Syncfusion.XlsIO.IWorksheet~ImportDataTable(DataTable,IName,Boolean,Int32,Int32).html">export data to a named range</a> from a specific row and column of the named range, you can make use of the below API, where rowOffset and columnOffset are the parameters to import from a particular cell in a named range.

```csharp
value = instance.ImportDataTable(dataTable, namedRange, showColumnName, rowOffset, colOffset);
```

<h2 id="Collections-to-Excel">2. Export from collection of objects to Excel</h2>

Exporting data from a collection of objects to an Excel worksheet is a common scenario. However, this option will be helpful if you need to export data from a model to an Excel worksheet.

The Syncfusion Excel (XlsIO) library provides support to export data from a collection of objects to an Excel worksheet.

Exporting data from a collection of objects to an Excel worksheet can be achieved through the ImportData method. The following code example shows how to export data from a collection to an Excel worksheet.

```csharp
using (ExcelEngine excelEngine = new ExcelEngine())
{
    IApplication application = excelEngine.Excel;
    application.DefaultVersion = ExcelVersion.Excel2016;

    //Read the data from XML file
    StreamReader reader = new StreamReader(Path.GetFullPath(@"../../Data/Customers.xml"));

    //Assign the data to the customerObjects collection
    IEnumerable customerObjects = GetData (reader.ReadToEnd());  

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
}
```

<img class="aligncenter wp-image-38681 size-full" src="https://blog.syncfusion.com/wp-content/uploads/2019/04/Collection-to-Excel.png" sizes="(max-width: 494px) 100vw, 494px" srcset="https://blog.syncfusion.com/wp-content/uploads/2019/04/Collection-to-Excel.png 494w, https://blog.syncfusion.com/wp-content/uploads/2019/04/Collection-to-Excel-300x291.png 300w" alt="Export collection of objects to Excel in c#" width="494" height="479" />

<em>Output of collection of objects to Excel</em>

<h2 id="Database-to-Excel">3. Export from Database to Excel</h2>

Excel supports creating Excel tables from different databases. If you have a scenario in which you need to create one or more Excel tables from a database using Excel, you need to establish every single connection to create those tables. This can be time consuming, so if you find an alternate way to generate Excel tables from database very quickly and easily, wouldn’t that be your first choice?

The Syncfusion Excel (XlsIO) library helps you to export data to Excel worksheets from databases like MS SQL, MS Access, Oracle, and more. By establishing a connection between the databases and Excel application, you can export data from a <strong>database to an Excel table</strong>.

You can use the <em><a href="https://help.syncfusion.com/cr/file-formats/Syncfusion.XlsIO.Base~Syncfusion.XlsIO.IListObject~Refresh.html">Refresh()</a></em><em> </em>option to update the modified data in the Excel table that is mapped to the database.

Above all, you can refer to the documentation to <u><a href="https://help.syncfusion.com/file-formats/xlsio/working-with-tables?cs-save-lang=1&amp;cs-lang=csharp#create-a-table-from-external-connection">create a table from an external connection</a></u> to learn more about how to export databases to Excel tables. The following code sample shows how to export data from a database to an Excel table.

```csharp
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
}
```

<img class="aligncenter wp-image-38692 size-full" src="https://blog.syncfusion.com/wp-content/uploads/2019/04/Database-to-Excel-1.png" sizes="(max-width: 736px) 100vw, 736px" srcset="https://blog.syncfusion.com/wp-content/uploads/2019/04/Database-to-Excel-1.png 736w, https://blog.syncfusion.com/wp-content/uploads/2019/04/Database-to-Excel-1-300x244.png 300w" alt="Export Database to Excel in c#" width="736" height="599" />

<em>Output of Database to Excel table</em>

<h2 id="DataGrid-GridView-DataGridView-to-Excel">4. Export data from DataGrid, GridView, DatGridView to Excel</h2>

Exporting data from <a href="https://docs.microsoft.com/en-us/dotnet/framework/winforms/controls/datagrid-control-overview-windows-forms">Microsoft grid</a> controls to Excel worksheets helps to visualize data in different ways. You may work for hours to iterate data and its styles from grid cells to export them into Excel worksheets. It should be good news for those who export data from Microsoft grid controls to Excel worksheets, because exporting with Syncfusion Excel library is much faster.

Syncfusion Excel (XlsIO) library supports to <a href="https://help.syncfusion.com/file-formats/xlsio/working-with-data#importing-data-from-microsoft-grid-controls-to-worksheet">exporting data from Microsoft Grid controls</a>, such as DataGrid, GridView, and DataGridView to Excel worksheets in a single API call. Also, you can export data with header and styles.

The following code example shows how to export data from DataGridView to an Excel worksheet.

```csharp
#region Loading the data to DataGridView
DataSet customersDataSet = new DataSet();

//Read the XML file with data
string inputXmlPath = Path.GetFullPath(@"../../Data/Employees.xml");
customersDataSet.ReadXml(inputXmlPath);
DataTable dataTable = new DataTable();
            
//Copy the structure and data of the table
dataTable = customersDataSet.Tables[1].Copy();
            
//Removing unwanted columns
dataTable.Columns.RemoveAt(0);
dataTable.Columns.RemoveAt(10);
this.dataGridView1.DataSource = dataTable;

dataGridView1.AlternatingRowsDefaultCellStyle.BackColor = Color.White;
dataGridView1.RowsDefaultCellStyle.BackColor = Color.LightBlue;
dataGridView1.ColumnHeadersDefaultCellStyle.Font = new System.Drawing.Font("Tahoma", 9F, ((System.Drawing.FontStyle)(System.Drawing.FontStyle.Bold)));
dataGridView1.ForeColor = Color.Black;
dataGridView1.BorderStyle = BorderStyle.None;
#endregion

using (ExcelEngine excelEngine = new ExcelEngine())
{
    IApplication application = excelEngine.Excel;

    //Create a workbook with single worksheet
    IWorkbook workbook = application.Workbooks.Create(1);

    IWorksheet worksheet = workbook.Worksheets[0];

    //Import from DataGridView to worksheet
    worksheet.ImportDataGridView(dataGridView1, 1, 1, isImportHeader: true, isImportStyle: true);

    worksheet.UsedRange.AutofitColumns();
    workbook.SaveAs("Output.xlsx");
}
```

<img class="aligncenter wp-image-38685 size-full" src="https://blog.syncfusion.com/wp-content/uploads/2019/04/Grid-to-Excel.png" sizes="(max-width: 970px) 100vw, 970px" srcset="https://blog.syncfusion.com/wp-content/uploads/2019/04/Grid-to-Excel.png 970w, https://blog.syncfusion.com/wp-content/uploads/2019/04/Grid-to-Excel-300x182.png 300w, https://blog.syncfusion.com/wp-content/uploads/2019/04/Grid-to-Excel-768x466.png 768w" alt="Export Microsoft DataGridView to Excel in c#" width="970" height="588" />

<em>Microsoft DataGridView to Excel</em>

<h2 id="Array-to-Excel">5. Export from array to Excel</h2>

Sometimes, there may be a need where an array of data may need to be inserted or modified into existing data in Excel worksheet. In this case, the number of rows and columns are known in advance. Arrays are useful when you have a fixed size.

The Syncfusion Excel (XlsIO) library provides support to export an array of data into an Excel worksheet, both horizontally and vertically. In addition, two-dimensional arrays can also be exported.

Let us consider a scenario, “Expenses per Person.” The expenses of a person for the whole year is tabulated in the Excel worksheet. In this scenario, you need to add expenses for a new person, <em>Paul Pogba,</em> in a new row and modify the expenses of all tracked people for the month <em>Dec</em>.

<img class="aligncenter wp-image-38677 size-full" src="https://blog.syncfusion.com/wp-content/uploads/2019/04/Array-to-Excel_1.png" sizes="(max-width: 982px) 100vw, 982px" srcset="https://blog.syncfusion.com/wp-content/uploads/2019/04/Array-to-Excel_1.png 982w, https://blog.syncfusion.com/wp-content/uploads/2019/04/Array-to-Excel_1-300x90.png 300w, https://blog.syncfusion.com/wp-content/uploads/2019/04/Array-to-Excel_1-768x231.png 768w" alt="Excel data before exporting from array to Excel in c#" width="982" height="296" />

<em>Excel data before exporting from array</em>

Exporting an array of data to Excel worksheet can be achieved through the <a href="https://help.syncfusion.com/file-formats/xlsio/working-with-data#import-data-from-array">ImportArray</a> method. The following code sample shows how to <strong>export an array of data to an Excel</strong> worksheet, both horizontally and vertically.

```csharp
using (ExcelEngine excelEngine = new ExcelEngine())
{
    IApplication application = excelEngine.Excel;
    application.DefaultVersion = ExcelVersion.Excel2016;

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
```

<img class="aligncenter wp-image-38678 size-full" src="https://blog.syncfusion.com/wp-content/uploads/2019/04/Array-to-Excel_2.png" sizes="(max-width: 976px) 100vw, 976px" srcset="https://blog.syncfusion.com/wp-content/uploads/2019/04/Array-to-Excel_2.png 976w, https://blog.syncfusion.com/wp-content/uploads/2019/04/Array-to-Excel_2-300x98.png 300w, https://blog.syncfusion.com/wp-content/uploads/2019/04/Array-to-Excel_2-768x250.png 768w" alt="Export array of data to Excel in c#" width="976" height="318" />

<em>Output of array of data to Excel</em>

<h2 id="CSV-to-Excel">6. Export from CSV to Excel</h2>

<a href="https://en.wikipedia.org/wiki/Comma-separated_values">Comma-separated value</a> (CSV) files are helpful in generating tabular data or lightweight reports with few columns and a high number of rows. Excel opens such files to make the data easier to read.

The Syncfusion Excel (XlsIO) library supports opening and saving CSV files in seconds. The below code example shows how to open a CSV file, also save it as XLSX file. Above all, the data is shown in a table with number formats applied.

```csharp
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
    table.BuiltInTableStyle =  TableBuiltInStyles.TableStyleMedium6;
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
```

<img class="aligncenter wp-image-38683 size-full" src="https://blog.syncfusion.com/wp-content/uploads/2019/04/CSV-to-Excel_Input.png" sizes="(max-width: 949px) 100vw, 949px" srcset="https://blog.syncfusion.com/wp-content/uploads/2019/04/CSV-to-Excel_Input.png 949w, https://blog.syncfusion.com/wp-content/uploads/2019/04/CSV-to-Excel_Input-300x182.png 300w, https://blog.syncfusion.com/wp-content/uploads/2019/04/CSV-to-Excel_Input-768x467.png 768w" alt="Input CSV File" width="949" height="577" />

<em>Input CSV file</em>

<img class="aligncenter wp-image-38684 size-full" src="https://blog.syncfusion.com/wp-content/uploads/2019/04/CSV-to-Excel_Output.png" sizes="(max-width: 1002px) 100vw, 1002px" srcset="https://blog.syncfusion.com/wp-content/uploads/2019/04/CSV-to-Excel_Output.png 1002w, https://blog.syncfusion.com/wp-content/uploads/2019/04/CSV-to-Excel_Output-300x156.png 300w, https://blog.syncfusion.com/wp-content/uploads/2019/04/CSV-to-Excel_Output-768x399.png 768w" alt="Export CSV to Excel in c#" width="1002" height="521" />

<em>Output of CSV converted to Excel</em>

Apart from this, <a href="https://www.syncfusion.com/excel-framework/net">Syncfusion Excel (XlsIO) library</a> provides various other features such as, charts, pivot tables, tables, cell formatting, conditional formatting, data validation, encryption and decryption, auto-shapes, Excel to PDF, Excel to Image, Excel to HTML and more. You can refer to our <a href="https://ej2.syncfusion.com/aspnetcore/XlsIO/Create#/material">online demo samples</a> to know more about all the features. 
