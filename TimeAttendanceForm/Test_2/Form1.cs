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
using Microsoft.WindowsAPICodePack.Dialogs;

namespace Test_2
{

    public partial class Form1 : Form
    {
        public static DateTime fileDatetime;

        public Form1()
        {
            InitializeComponent();
            openFileDialog1 = new OpenFileDialog();
        }

        private void Switch_Gb_Project(bool Data,bool Cell)
        {
            Panel_Groupbox_Data_Project.Enabled = Data;
            Panel_Groupbox_Cell_Project.Enabled = Cell;
        }

        private void Switch_Gb_Site_Start(bool Data, bool Cell)
        {
            Panel_Groupbox_Data_Site_Start.Enabled = Data;
            Panel_Groupbox_Cell_Site_Start.Enabled = Cell;
        }

        private void Switch_Gb_Site_Stop(bool Data, bool Cell)
        {
            Panel_Groupbox_Data_Site_Stop.Enabled = Data;
            Panel_Groupbox_Cell_Site_Stop.Enabled = Cell;
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
            try
            {
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
                                                                , (getIntFromAddr(tb_Date.Text) - 1)
                                                                , getIntFromAddr(tb_Date.Text)
                                                                , GetColumnNumber(getStringFromAddr(tb_Date.Text))
                                                                , GetColumnNumber(getStringFromAddr(tb_Time_In.Text))
                                                                , GetColumnNumber(getStringFromAddr(tb_Time_Out.Text))
                                                                , tb_Emp_No.Text
                                                                , tb_Name.Text);
                        Dgv_Show_Preview.DataSource = tbContainer;
                        Dgv_Show_Preview.Columns[3].Visible = false; // hide raw date
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
            catch
            {
                MessageBox.Show("Please check inputbox.");
            }
               
        }
        public bool checkSameRow()
        {
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
        public string getIntOnly(string str)
        {
            Regex regex = new Regex(@"^\d$");
            var match = regex.Match(str);

            return match.Groups[0].Value;
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
        
        public DataTable LoadExcelFile_specific_colunm(string fileName, string worksheetName, int headerRowNumber, int firstDataRowNumber, int colDateTime, int colTimeIn, int colTimeOut,string colEmpNo,string colEmpName)
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
            fileDatetime = DateTime.FromOADate(edate); 

            int dayMonth = System.DateTime.DaysInMonth(fileDatetime.Year, fileDatetime.Month);

            //----------------------------- end date in month--------------------------------------
            //----------------------------- get emp no --------------------------------------------
            lb_Emp_No.Text = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)ExcelWorksheet.Cells[ getIntFromAddr(colEmpNo), GetColumnNumber( getStringFromAddr(colEmpNo))]).Value2);
            lb_Emp_No.Text = Regex.Match(lb_Emp_No.Text, @"\d+").Value;
            //----------------------------- end emp no --------------------------------------------
            //----------------------------- get emp name --------------------------------------------
            lb_Name.Text = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)ExcelWorksheet.Cells[getIntFromAddr(colEmpName), GetColumnNumber(getStringFromAddr(colEmpName))]).Value2);
            //----------------------------- end emp name --------------------------------------------

            ExcelWorksheet = (Microsoft.Office.Interop.Excel.Worksheet)ExcelWorkbook.Worksheets[WorksheetName];
           
            dt.Columns.Add("Date");
            dt.Columns.Add("Time IN");
            dt.Columns.Add("Time Out");
            dt.Columns.Add("Date Raw");


            for (int i = 0; i < dayMonth; i++)
            {
                try
                {
                 

                    DataRow Row = dt.NewRow();

                    string sDate;

                    double date;
                    try
                    {
                        sDate = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)ExcelWorksheet.Cells[i + firstDataRowNumber, colDateTime]).Value2);

                        date = double.Parse(sDate);
                    
                        Row[0] = DateTime.FromOADate(date).ToString("dd/MM/yyyy");
                        Row[3] = date;
                    }
                    catch
                    {
                        sDate = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)ExcelWorksheet.Cells[i + firstDataRowNumber, colDateTime]).Value2);

                        Row[0] = sDate;
                        Row[3] = sDate;
                    }
                    try
                    {
                        sDate = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)ExcelWorksheet.Cells[i + firstDataRowNumber, colTimeIn]).Value2);

                        date = double.Parse(sDate);

                        Row[1] = DateTime.FromOADate(date).ToString("HH:mm");

                        sDate = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)ExcelWorksheet.Cells[i + firstDataRowNumber, colTimeOut]).Value2);

                        date = double.Parse(sDate);

                        Row[2] = DateTime.FromOADate(date).ToString("HH:mm");
                    }
                    catch
                    {
                    }


                    if (!string.IsNullOrWhiteSpace(Row[0].ToString()))
                    {
                        dt.Rows.Add(Row); 
                    }


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

        private void btn_Browse_Dest_file_Click(object sender, EventArgs e)
        {
            var Selected_folder = new CommonOpenFileDialog();
            Selected_folder.IsFolderPicker = true;
            CommonFileDialogResult result = Selected_folder.ShowDialog();
            if (result == CommonFileDialogResult.Ok)
            {
                tb_Dest_file.Text = Selected_folder.FileName;
            }
        }

        private void Export_Excell()
        {
            try
            {
                Excel.Application excelApp = new Excel.Application();
                Excel.Workbook workBook = excelApp.Workbooks.Open(openFileDialog2.FileName);
                string WorksheetName = workBook.ActiveSheet.Name;
                Excel.Worksheet workSheet = (Excel.Worksheet)workBook.Worksheets[WorksheetName];


                for (int i = 0; i < Dgv_Show_Preview.Rows.Count; i++) // i = row
                {

                    workSheet.Cells[i + 2, GetColumnNumber(getStringFromAddr("A2"))] = lb_Emp_No.Text;
                    workSheet.Cells[i + 2, GetColumnNumber(getStringFromAddr("B2"))] = lb_Name.Text;
                    //for (int j = 0; j < Dgv_Show_Preview.Columns.Count; j++) //j = column
                    //{
                    //    if (Dgv_Show_Preview.Rows[i].Cells[j].Value != null)
                    //        workSheet.Cells[i + 2, j + 3] = Dgv_Show_Preview.Rows[i].Cells[j].Value.ToString();
                    //    else
                    //        workSheet.Cells[i + 2, j + 3] = "";
                    //}

                    workSheet.Cells[i + 2, GetColumnNumber(getStringFromAddr("C2"))].EntireColumn.NumberFormat = "dd/MM/yyyy";
                    workSheet.Cells[i + 2, GetColumnNumber(getStringFromAddr("C2"))] = Dgv_Show_Preview.Rows[i].Cells[3].Value.ToString();
                    workSheet.Cells[i + 2, GetColumnNumber(getStringFromAddr("D2"))] = Dgv_Show_Preview.Rows[i].Cells[1].Value.ToString();
                    workSheet.Cells[i + 2, GetColumnNumber(getStringFromAddr("E2"))] = Dgv_Show_Preview.Rows[i].Cells[2].Value.ToString();
                    workSheet.Cells[i + 2, GetColumnNumber(getStringFromAddr("F2"))] = tb_Groupbox_Data_Site_Start.Text;
                    workSheet.Cells[i + 2, GetColumnNumber(getStringFromAddr("G2"))] = tb_Groupbox_Data_Site_Stop.Text;
                    workSheet.Cells[i + 2, GetColumnNumber(getStringFromAddr("H2"))] = tb_Groupbox_Data_Project.Text;
                }

                workBook.SaveAs(tb_Dest_file.Text + "\\" + tb_Groupbox_Data_Project.Text + "_" + fileDatetime.ToString("MMMyyyy") + ".xls");  // NOTE: You can use 'Save()' or 'SaveAs()'
                workBook.Close();
                excelApp.Quit();
                Console.WriteLine(tb_Dest_file.Text + "\\" + tb_Groupbox_Data_Project.Text + "_" + fileDatetime.ToString("MMMyyyy") + ".xls");
            }
            catch(Exception ex) {
                Console.WriteLine(ex.StackTrace);
            }
        }

        private void btn_Export_Click(object sender, EventArgs e)
        {
            if (Dgv_Show_Preview.Rows.Count == 0)
            {
                btn_Preview_Click(null, e);
            }
            Export_Excell();
        }

        private void btn_Browse_Template_file_Click(object sender, EventArgs e)
        {
            
            openFileDialog2.Filter = "XML Files (*.xml; *.xls; *.xlsx; *.xlsm; *.xlsb) |*.xml; *.xls; *.xlsx; *.xlsm; *.xlsb";//open file format define Excel Files(.xls)|*.xls| Excel Files(.xlsx)|*.xlsx| 
            openFileDialog2.FilterIndex = 3;

            openFileDialog2.Multiselect = false;        //not allow multiline selection at the file selection level
            openFileDialog2.Title = "Select file to import";   //define the name of openfileDialog
            openFileDialog2.InitialDirectory = @"Desktop"; //define the initial directory
            DialogResult result = openFileDialog2.ShowDialog(); // Show the dialog.
            if (result == DialogResult.OK) // Test result.
            {
                string Pathfile = openFileDialog2.FileName;
                try
                {
                    tb_Template_file.Text = @"...\" + Path.GetFileName(Pathfile);
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.StackTrace);
                }
            }
        }

        private void Panel_Groupbox_Data_Project_MouseClick(object sender, MouseEventArgs e)
        {
            Switch_Gb_Project(true, false);
        }

        private void Panel_Groupbox_Cell_Project_MouseClick(object sender, MouseEventArgs e)
        {
            Switch_Gb_Project(false, true);
        }

        private void Panel_Groupbox_Data_Site_Start_MouseClick(object sender, MouseEventArgs e)
        {
            Switch_Gb_Site_Start(true, false);
        }

        private void Panel_Groupbox_Cell_Site_Start_MouseClick(object sender, MouseEventArgs e)
        {
            Switch_Gb_Site_Start(false, true);
        }

        private void Panel_Groupbox_Data_Site_Stop_MouseClick(object sender, MouseEventArgs e)
        {
            Switch_Gb_Site_Stop(true, false);
        }

        private void Panel_Groupbox_Cell_Site_Stop_MouseClick(object sender, MouseEventArgs e)
        {
            Switch_Gb_Site_Stop(false, true);
        }
    }// end form1
}
