using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace Test_2
{
    public partial class Form1 : Form
    {
        //public OpenFileDialog openFileDialog1;
        public Form1()
        {
            InitializeComponent();
            openFileDialog1 = new OpenFileDialog();
        }

        private void Switch_Groupbox(bool data,bool cell)
        // frist value is groupbox data
        // second value is groupbox cell
        {
            gb_Data.Enabled = data;
            gb_Cell.Enabled = cell;
        }
   
        private void Panel_Groupbox_Cell_MouseClick(object sender, MouseEventArgs e)
        {
            Switch_Groupbox(false, true);
        }

        private void Panel_Groupbox_Data_MouseClick(object sender, MouseEventArgs e)
        {
            Switch_Groupbox(true, false);
        }

        private void btn_Browse_Target_file_Click(object sender, EventArgs e)
        {
            
            openFileDialog1.Filter = "XML Files (*.xml; *.xls; *.xlsx; *.xlsm; *.xlsb) |*.xml; *.xls; *.xlsx; *.xlsm; *.xlsb";//open file format define Excel Files(.xls)|*.xls| Excel Files(.xlsx)|*.xlsx| 
            openFileDialog1.FilterIndex = 3;

            openFileDialog1.Multiselect = false;        //not allow multiline selection at the file selection level
            openFileDialog1.Title = "Select file to import";   //define the name of openfileDialog
            openFileDialog1.InitialDirectory = @"Desktop"; //define the initial directory
            DialogResult result = openFileDialog1.ShowDialog(); // Show the dialog.
            if (result == DialogResult.OK) // Test result.
            {
                string Pathfile = openFileDialog1.FileName;
                try
                {
                    tb_Target_file.Text = @"...\" + Path.GetFileName(Pathfile);
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.StackTrace);
                }
            }
        }

        private void btn_Preview_Click(object sender, EventArgs e)
        {
            Console.WriteLine(openFileDialog1.FileName);
            if (!string.IsNullOrWhiteSpace(openFileDialog1.FileName))
            {
                string pathName = this.openFileDialog1.FileName;
                string fileName = System.IO.Path.GetFileNameWithoutExtension(this.openFileDialog1.FileName);
                DataTable tbContainer = new DataTable();
                string strConn = string.Empty;
                string sheetName = "";
                if (checkSameRow())
                {
                    tbContainer = LoadExcelFile_specific_colunm(pathName
                                                            , sheetName
                                                            , (getIntFromAddr(tb_Date .Text)-1)
                                                            , getIntFromAddr(tb_Date.Text)
                                                            , GetColumnNumber(getStringFromAddr(tb_Date.Text))
                                                            , GetColumnNumber(getStringFromAddr(tb_Time_In.Text))
                                                            , GetColumnNumber(getStringFromAddr(tb_Time_Out.Text)));
                    Dgv_Show_Preview.DataSource = tbContainer;
                }
                else
                {
                    MessageBox.Show("Error! Check row Date, Time in, Time out");
                }
            }
            else
            {
                MessageBox.Show("Error! Please select file to import.");
            }
               
        }
        public bool checkSameRow()
        {
            Console.WriteLine(getIntFromAddr(tb_Date.Text));
            if (getIntFromAddr(tb_Date.Text) == getIntFromAddr(tb_Time_In.Text)
                && getIntFromAddr(tb_Date.Text) == getIntFromAddr(tb_Time_Out.Text))
            {
                return true;
            }
            else
            {
                return false;
            }
        }
        public string getStringFromAddr(string str)
        {
            var numAlpha = new Regex("(?<Alpha>[a-zA-Z]*)(?<Numeric>[0-9]*)");
            var match = numAlpha.Match(str);

            return match.Groups["Alpha"].Value;
        }
        public int getIntFromAddr(string str)
        {
            var numAlpha = new Regex("(?<Alpha>[a-zA-Z]*)(?<Numeric>[0-9]*)");
            var match = numAlpha.Match(str);

            return Int32.Parse(match.Groups["Numeric"].Value);
        }
        public static int GetColumnNumber(string name)
        {
            int number = 0;
            int pow = 1;
            for (int i = name.Length - 1; i >= 0; i--)
            {
                number += (name.ToUpper()[i] - 'A' + 1) * pow;
                pow *= 26;
            }

            return number;
        }
        public string GetExcelColumnName(int columnNumber)
        {
            int dividend = columnNumber;
            string columnName = String.Empty;
            int modulo;

            while (dividend > 0)
            {
                modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo).ToString() + columnName;
                dividend = (int)((dividend - modulo) / 26);
            }

            return columnName;
        }
        public static DataTable LoadExcelFile_specific_colunm(string fileName, string worksheetName, int headerRowNumber, int firstDataRowNumber, int colDateTime, int colTimeIn, int colTimeOut)
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
            //----------------------------- get date in month--------------------------------------
            string eeDate = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)ExcelWorksheet.Cells[firstDataRowNumber, colDateTime]).Value2);
            //Console.WriteLine(eeDate);
            double edate = double.Parse(eeDate);
            DateTime datetime = DateTime.FromOADate(edate);

            int dayMonth = System.DateTime.DaysInMonth(datetime.Year, datetime.Month);

            //----------------------------- end date in month--------------------------------------

            ExcelWorksheet = (Microsoft.Office.Interop.Excel.Worksheet)ExcelWorkbook.Worksheets[WorksheetName];
            for (int i = 0; i < ExcelWorksheet.UsedRange.Columns.Count; i++)
            {
                string ColumnHeading = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)ExcelWorksheet.Cells[headerRowNumber, i + 1]).Value2);

                if (!String.IsNullOrWhiteSpace(ColumnHeading) && !dt.Columns.Contains(ColumnHeading))
                {

                    if (Columns.ContainsKey(ColumnHeading))
                    {
                        Columns.Add(ColumnHeading + "2", i + 1);
                    }
                    else
                    {
                        Columns.Add(ColumnHeading, i + 1);
                    }

                    if (i + 1 == colDateTime || i + 1 == colTimeIn || i + 1 == colTimeOut)
                    {
                        dt.Columns.Add(ColumnHeading);
                    }

                }


            }
            for (int i = 0; i < dayMonth; i++)
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
                            try
                            {
                                string sDate = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)ExcelWorksheet.Cells[i + firstDataRowNumber, kvp.Value]).Value2);

                                double date = double.Parse(sDate);

                                CellContent = DateTime.FromOADate(date).ToString("dd/MM/yyyy");
                            }
                            catch
                            {
                                CellContent = "";
                            }
                            Row[ColumnCount] = CellContent;
                            ColumnCount++;

                        }
                        else if (kvp.Value == colTimeIn || kvp.Value == colTimeOut)
                        {
                            try
                            {
                                string sDate = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)ExcelWorksheet.Cells[i + firstDataRowNumber, kvp.Value]).Value2);

                                double date = double.Parse(sDate);

                                CellContent = DateTime.FromOADate(date).ToString("HH:mm");
                            }
                            catch
                            {
                                CellContent = " ";
                            }
                            Row[ColumnCount] = CellContent;
                            ColumnCount++;
                        }





                        if (!string.IsNullOrWhiteSpace(CellContent))
                        {
                            RowHasContent = true;

                        }

                    }
                    Console.Write(Row[0].ToString());
                    if (RowHasContent)
                    {
                        dt.Rows.Add(Row); ;

                    }
                    Console.WriteLine(i);
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.StackTrace);

                }

            }

            // Clean up

            try { ExcelWorksheet = null; } catch { }

            try { ExcelWorkbook.Close(); } catch { }

            try { ExcelWorkbook = null; } catch { }

            try { ExcelApplication = null; } catch { }
            return dt;
        }
    }// end form1
}
