using Syncfusion. XlsIO;
using System;
using System. Collections. Generic;
using System. ComponentModel;
using System. Data;
using System. Drawing;
using System. IO;
using System. Linq;
using System. Text;
using System. Threading. Tasks;
using System. Windows. Forms;

namespace ImportFromGrid
{
    public partial class Form1 :Form
    {
        public Form1()
        {
            InitializeComponent();

            #region Loading the data to Data Grid
            DataSet customersDataSet = new DataSet();
            //Get the path of the input file
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


        }

        private void btnCreate_Click(object sender, System.EventArgs e)
        {

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
                System.Diagnostics.Process.Start("Output.xlsx");
            }
        }
    }
}
