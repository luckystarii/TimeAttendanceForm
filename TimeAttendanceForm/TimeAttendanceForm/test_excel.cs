using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

using System;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using System.Data.OleDb;
using System.IO;

namespace TimeAttendanceForm
{
    public partial class Form2 : Form
    {
        public Form2()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {

            try
            {
                OpenFileDialog openFileDialog1 = new OpenFileDialog();  //create openfileDialog Object
                openFileDialog1.Filter = "XML Files (*.xml; *.xls; *.xlsx; *.xlsm; *.xlsb) |*.xml; *.xls; *.xlsx; *.xlsm; *.xlsb";//open file format define Excel Files(.xls)|*.xls| Excel Files(.xlsx)|*.xlsx| 
                openFileDialog1.FilterIndex = 3;

                openFileDialog1.Multiselect = false;        //not allow multiline selection at the file selection level
                openFileDialog1.Title = "Open Text File-R13";   //define the name of openfileDialog
                openFileDialog1.InitialDirectory = @"Desktop"; //define the initial directory

                if (openFileDialog1.ShowDialog() == DialogResult.OK)        //executing when file open
                {
                    string pathName = openFileDialog1.FileName;
                    string fileName = System.IO.Path.GetFileNameWithoutExtension(openFileDialog1.FileName);
                    DataTable tbContainer = new DataTable();
                    string strConn = string.Empty;
                    string sheetName = "";

                    //FileInfo file = new FileInfo(pathName);
                    //if (!file.Exists) { throw new Exception("Error, file doesn't exists!"); }
                    //string extension = file.Extension;
                    //switch (extension)
                    //{
                    //    case ".xls":
                    //        strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + pathName + ";Extended Properties='Excel 8.0;HDR=Yes;IMEX=1;'";
                    //        break;
                    //    case ".xlsx":
                    //        strConn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + pathName + ";Extended Properties='Excel 12.0;HDR=Yes;IMEX=1;'";
                    //        break;
                    //    default:
                    //        strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + pathName + ";Extended Properties='Excel 8.0;HDR=Yes;IMEX=1;'";
                    //        break;
                    //}
                    //OleDbConnection cnnxls = new OleDbConnection(strConn);
                    //OleDbDataAdapter oda = new OleDbDataAdapter(string.Format("select * from [{0}$]", sheetName), cnnxls);
                    //oda.Fill(tbContainer);

                    //textBox1.Text = tbContainer.Rows[4][0].ToString();
                    //dataGridView1.DataSource = tbContainer;

                    tbContainer = LoadExcelFile_specific_colunm(pathName, sheetName, 4, 4,2,5,7);
                    dataGridView1.DataSource = tbContainer;
                }

            }
            catch (Exception)
            {
                MessageBox.Show("Error!");
            }

        }// end function button1_Click

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        public static DataTable LoadExcelFile(string fileName, string worksheetName, int headerRowNumber, int firstDataRowNumber)
        {
            
            DataTable dt = new DataTable();

            Microsoft.Office.Interop.Excel.Application ExcelApplication = new Microsoft.Office.Interop.Excel.Application();

            Microsoft.Office.Interop.Excel.Workbook ExcelWorkbook = ExcelApplication.Workbooks.Open(fileName, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);

            Microsoft.Office.Interop.Excel.Worksheet ExcelWorksheet = null;

            string WorksheetName = worksheetName;

            if (string.IsNullOrWhiteSpace(worksheetName))
            {
                WorksheetName = ExcelWorkbook.ActiveSheet.Name;

            }

            ExcelWorksheet = (Microsoft.Office.Interop.Excel.Worksheet)ExcelWorkbook.Worksheets[WorksheetName];

            dt.TableName = WorksheetName;

            // Add the columns

            Dictionary<string, int> Columns = new Dictionary<string, int>();

            for (int i = 0; i < ExcelWorksheet.UsedRange.Columns.Count; i++)
            {
                string ColumnHeading = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)ExcelWorksheet.Cells[headerRowNumber, i + 1]).Value2);

                if (!String.IsNullOrWhiteSpace(ColumnHeading) && !dt.Columns.Contains(ColumnHeading))
                {
                    Columns.Add(ColumnHeading, i + 1);

                    dt.Columns.Add(ColumnHeading);

                }

            }

            // Add the rows

            for (int i = 0; i < ExcelWorksheet.UsedRange.Rows.Count - firstDataRowNumber + 1; i++)
            {
                try
                {
                    int ColumnCount = 0;

                    DataRow Row = dt.NewRow();

                    bool RowHasContent = false;

                    foreach (KeyValuePair<string, int> kvp in Columns)
                    {
                        string CellContent = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)ExcelWorksheet.Cells[i + firstDataRowNumber, kvp.Value]).Value2);

                        Row[ColumnCount] = CellContent;

                        ColumnCount++;

                        if (!string.IsNullOrWhiteSpace(CellContent))
                        {
                            RowHasContent = true;

                        }

                    }

                    if (RowHasContent)
                    {
                        dt.Rows.Add(Row); ;

                    }

                }
                catch
                {

                }

            }

            // Clean up

            try { ExcelWorksheet = null; } catch { }

            try { ExcelWorkbook.Close(); } catch { }

            try { ExcelWorkbook = null; } catch { }

            try { ExcelApplication = null; } catch { }

            return dt;
        }

        public static DataTable LoadExcelFile_specific_colunm(string fileName, string worksheetName, int headerRowNumber, int firstDataRowNumber,int colDateTime,int colTimeIn,int colTimeOut)
        {
            DataTable dt = new DataTable();

            Microsoft.Office.Interop.Excel.Application ExcelApplication = new Microsoft.Office.Interop.Excel.Application();

            Microsoft.Office.Interop.Excel.Workbook ExcelWorkbook = ExcelApplication.Workbooks.Open(fileName, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);

            Microsoft.Office.Interop.Excel.Worksheet ExcelWorksheet = null;

            string WorksheetName = worksheetName;
            
            if (string.IsNullOrWhiteSpace(worksheetName))
            {
                WorksheetName = ExcelWorkbook.ActiveSheet.Name;

            }

            ExcelWorksheet = (Microsoft.Office.Interop.Excel.Worksheet)ExcelWorkbook.Worksheets[WorksheetName];

            dt.TableName = WorksheetName;

            Dictionary<string, int> Columns = new Dictionary<string, int>();

            for (int i = 0; i < ExcelWorksheet.UsedRange.Columns.Count; i++)
            {
                string ColumnHeading = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)ExcelWorksheet.Cells[headerRowNumber, i + 1]).Value2);

                if (!String.IsNullOrWhiteSpace(ColumnHeading) && !dt.Columns.Contains(ColumnHeading))
                {
                    Columns.Add(ColumnHeading, i + 1);

                    dt.Columns.Add(ColumnHeading);

                }

            }
              for (int i = 0; i < ExcelWorksheet.UsedRange.Rows.Count - firstDataRowNumber + 1; i++)
            {
                try
                {
                    int ColumnCount = 0;

                    DataRow Row = dt.NewRow();

                    bool RowHasContent = false;

                    foreach (KeyValuePair<string, int> kvp in Columns)
                    {
                        //string data = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)ExcelWorksheet.Cells[i + firstDataRowNumber, kvp.Value]).Value2);
                        string CellContent = null;
                        if (kvp.Value == colDateTime)
                        {
                             //CellContent = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)ExcelWorksheet.Cells[i + firstDataRowNumber, kvp.Value]).Value2);
                             string sDate = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)ExcelWorksheet.Cells[i + firstDataRowNumber, kvp.Value]).Value2);
                             
                             double date = double.Parse(sDate);

                             CellContent = DateTime.FromOADate(date).ToString("dd/MMM/yyyy");
                        }
                        else if (kvp.Value == colTimeIn || kvp.Value == colTimeOut)
                        {
                            string sDate = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)ExcelWorksheet.Cells[i + firstDataRowNumber, kvp.Value]).Value2);

                            double date = double.Parse(sDate);

                            CellContent = DateTime.FromOADate(date).ToString("hh:mm");
                        }
                        else
                        {
                             CellContent = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)ExcelWorksheet.Cells[i + firstDataRowNumber, kvp.Value]).Value2);

                        }

                        Row[ColumnCount] = CellContent;
                        ColumnCount++;

                        if (!string.IsNullOrWhiteSpace(CellContent))
                        {
                            RowHasContent = true;

                        }

                    }

                    if (RowHasContent)
                    {
                        dt.Rows.Add(Row); ;

                    }

                }
                catch
                {

                }

            }

            // Clean up

            try { ExcelWorksheet = null; } catch { }

            try { ExcelWorkbook.Close(); } catch { }

            try { ExcelWorkbook = null; } catch { }

            try { ExcelApplication = null; } catch { }
            return dt;
        }
    }
}
