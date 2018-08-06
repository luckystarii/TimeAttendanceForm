﻿using Microsoft.WindowsAPICodePack.Dialogs;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace TAIE
{

    public partial class Home : Form
    {
        public static DateTime fileDatetime;

        public static String ProjectName;
        public static String SiteStartTime;
        public static String SiteStopTime;

        public Home()
        {
            InitializeComponent();
            Default_Data_AND_Cell();
            openFileDialog1 = new OpenFileDialog();

        }

        private void Default_Data_AND_Cell()
        {
            tb_Groupbox_Cell_Project.Enabled = false;
            tb_Groupbox_Cell_Site_Start.Enabled = false;
            tb_Groupbox_Cell_Site_Stop.Enabled = false;
        }

        private void Switch_Gb_Project(bool Data, bool Cell)
        {
            tb_Groupbox_Data_Project.Enabled = Data;
            tb_Groupbox_Cell_Project.Enabled = Cell;
        }

        private void Switch_Gb_Site_Start(bool Data, bool Cell)
        {
            tb_Groupbox_Data_Site_Start.Enabled = Data;
            tb_Groupbox_Cell_Site_Start.Enabled = Cell;
        }

        private void Switch_Gb_Site_Stop(bool Data, bool Cell)
        {
            tb_Groupbox_Data_Site_Stop.Enabled = Data;
            tb_Groupbox_Cell_Site_Stop.Enabled = Cell;
        }

        private void All_Clear_Data()
        {
            tb_Target_file.Text = "";
            tb_Template_file.Text = "";
            tb_Dest_file.Text = "";
            tb_Groupbox_Data_Project.Text = "";
            tb_Groupbox_Cell_Project.Text = "";
            tb_Groupbox_Data_Site_Start.Text = "";
            tb_Groupbox_Cell_Site_Start.Text = "";
            tb_Groupbox_Data_Site_Stop.Text = "";
            tb_Groupbox_Cell_Site_Stop.Text = "";
            tb_Emp_No.Text = "";
            tb_Name.Text = "";
            tb_Date.Text = "";
            tb_Time_In.Text = "";
            tb_Time_Out.Text = "";
        }
        private bool Check_EmptyValues()
        {
            if (string.IsNullOrEmpty(tb_Target_file.Text.Trim())
                && string.IsNullOrEmpty(tb_Template_file.Text.Trim())
                && string.IsNullOrEmpty(tb_Dest_file.Text.Trim())
                && (string.IsNullOrEmpty(tb_Groupbox_Data_Project.Text.Trim()) || !tb_Groupbox_Data_Project.Enabled)
                && (string.IsNullOrEmpty(tb_Groupbox_Cell_Project.Text.Trim()) || !tb_Groupbox_Cell_Project.Enabled)
                && (string.IsNullOrEmpty(tb_Groupbox_Data_Site_Start.Text.Trim()) || !tb_Groupbox_Data_Site_Start.Enabled)
                && (string.IsNullOrEmpty(tb_Groupbox_Cell_Site_Start.Text.Trim()) || !tb_Groupbox_Cell_Site_Start.Enabled)
                && (string.IsNullOrEmpty(tb_Groupbox_Data_Site_Stop.Text.Trim()) || !tb_Groupbox_Data_Site_Stop.Enabled)
                && (string.IsNullOrEmpty(tb_Groupbox_Cell_Site_Stop.Text.Trim()) || !tb_Groupbox_Cell_Site_Stop.Enabled)
                && string.IsNullOrEmpty(tb_Emp_No.Text.Trim())
                && string.IsNullOrEmpty(tb_Name.Text.Trim()) 
                && string.IsNullOrEmpty(tb_Date.Text.Trim())
                && string.IsNullOrEmpty(tb_Time_In.Text.Trim())
                && string.IsNullOrEmpty(tb_Time_Out.Text.Trim()))
            {
                return true;
            }
            return false;
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
                    if (checkSameRow(tb_Date.Text,tb_Time_In.Text,tb_Time_Out.Text))
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
        public bool checkSameRow(string str1,string str2,string str3)
        {
            if (getIntFromAddr(str1) == getIntFromAddr(str2)
                && getIntFromAddr(str1) == getIntFromAddr(str3))
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

            //----------------------------- get project Name ----------------------------------------
            if (tb_Groupbox_Data_Project.Enabled == true)
            {
                ProjectName = tb_Groupbox_Data_Project.Text;
            }
            else
            {
                ProjectName = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)ExcelWorksheet.Cells[getIntFromAddr(tb_Groupbox_Cell_Project.Text), GetColumnNumber(getStringFromAddr(tb_Groupbox_Cell_Project.Text))]).Value2);
            }
            //----------------------------- end project Name ----------------------------------------


            //----------------------------- get site start time ----------------------------------------
            if (tb_Groupbox_Data_Site_Start.Enabled == true)
            {
                SiteStartTime = tb_Groupbox_Data_Site_Start.Text;
            }
            else
            {
                SiteStartTime = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)ExcelWorksheet.Cells[getIntFromAddr(tb_Groupbox_Cell_Site_Start.Text), GetColumnNumber(getStringFromAddr(tb_Groupbox_Cell_Site_Start.Text))]).Value2);
            }
            //----------------------------- end site start time ----------------------------------------

            //----------------------------- get site stop time ----------------------------------------
            if (tb_Groupbox_Data_Site_Stop.Enabled == true)
            {
                SiteStopTime = tb_Groupbox_Data_Site_Stop.Text;
            }
            else
            {
                SiteStopTime = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)ExcelWorksheet.Cells[getIntFromAddr(tb_Groupbox_Cell_Site_Stop.Text), GetColumnNumber(getStringFromAddr(tb_Groupbox_Cell_Site_Stop.Text))]).Value2);
            }
            //----------------------------- end site stop time ----------------------------------------

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

                if (checkSameRow(export_date.Text, export_time_in.Text, export_time_out.Text))
                {
                    for (int i = 0; i < Dgv_Show_Preview.Rows.Count; i++) // i = row
                    {

                        workSheet.Cells[getIntFromAddr(export_time_out.Text) + i, GetColumnNumber(getStringFromAddr(export_emp_no.Text))] = lb_Emp_No.Text;
                        workSheet.Cells[getIntFromAddr(export_time_out.Text) + i, GetColumnNumber(getStringFromAddr(export_emp_name.Text))] = lb_Name.Text;
                        //for (int j = 0; j < Dgv_Show_Preview.Columns.Count; j++) //j = column
                        //{
                        //    if (Dgv_Show_Preview.Rows[i].Cells[j].Value != null)
                        //        workSheet.Cells[i + 2, j + 3] = Dgv_Show_Preview.Rows[i].Cells[j].Value.ToString();
                        //    else
                        //        workSheet.Cells[i + 2, j + 3] = "";
                        //}

                        workSheet.Cells[getIntFromAddr(export_time_out.Text)+i, GetColumnNumber(getStringFromAddr(export_date.Text))].EntireColumn.NumberFormat = "dd/MM/yyyy";
                        workSheet.Cells[getIntFromAddr(export_time_out.Text)+i, GetColumnNumber(getStringFromAddr(export_time_in.Text))] = Dgv_Show_Preview.Rows[i].Cells[1].Value.ToString();
                        workSheet.Cells[getIntFromAddr(export_time_out.Text)+i, GetColumnNumber(getStringFromAddr(export_time_out.Text))] = Dgv_Show_Preview.Rows[i].Cells[2].Value.ToString();
                        workSheet.Cells[getIntFromAddr(export_time_out.Text)+i, GetColumnNumber(getStringFromAddr(export_project_name.Text))] = ProjectName;
                        workSheet.Cells[getIntFromAddr(export_time_out.Text)+i, GetColumnNumber(getStringFromAddr(export_date.Text))] = Dgv_Show_Preview.Rows[i].Cells[3].Value.ToString();
                        workSheet.Cells[getIntFromAddr(export_time_out.Text)+i, GetColumnNumber(getStringFromAddr(export_site_start.Text))] = SiteStartTime;
                        workSheet.Cells[getIntFromAddr(export_time_out.Text) + i, GetColumnNumber(getStringFromAddr(export_site_stop.Text))] = SiteStopTime;
                    }
                }
                else
                {
                    MessageBox.Show("Error! Check row in Config tab.");
                }
                workBook.SaveAs(tb_Dest_file.Text + "\\" + tb_Groupbox_Data_Project.Text + "_" + lb_Name.Text + "_" + fileDatetime.ToString("MMMyyyy") + ".xls");  // NOTE: You can use 'Save()' or 'SaveAs()'
                workBook.Close();
                excelApp.Quit();
                MessageBox.Show("Export Complete. \n File : "+ tb_Dest_file.Text + "\\" + tb_Groupbox_Data_Project.Text + "_" + lb_Name.Text + "_" + fileDatetime.ToString("MMMyyyy") + ".xls");
            }
            catch(Exception ex)
            {
                MessageBox.Show("Export Error! Please check Project Detail.");
            }
        }

        private void btn_Export_Click(object sender, EventArgs e)
        {
            if (Dgv_Show_Preview.Rows.Count == 0)
            {
                btn_Preview_Click(null, e);
            }
            if (Dgv_Show_Preview.Rows.Count != 0)
            {
                Export_Excell();
            }
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
            tb_Groupbox_Data_Project.Focus();
        }

        private void Panel_Groupbox_Cell_Project_MouseClick(object sender, MouseEventArgs e)
        {
            Switch_Gb_Project(false, true);
            tb_Groupbox_Cell_Project.Focus();
        }

        private void Panel_Groupbox_Data_Site_Start_MouseClick(object sender, MouseEventArgs e)
        {
            Switch_Gb_Site_Start(true, false);
            tb_Groupbox_Data_Site_Start.Focus();
        }

        private void Panel_Groupbox_Cell_Site_Start_MouseClick(object sender, MouseEventArgs e)
        {
            Switch_Gb_Site_Start(false, true);
            tb_Groupbox_Cell_Site_Start.Focus();
        }

        private void Panel_Groupbox_Data_Site_Stop_MouseClick(object sender, MouseEventArgs e)
        {
            Switch_Gb_Site_Stop(true, false);
            tb_Groupbox_Data_Site_Stop.Focus();
        }

        private void Panel_Groupbox_Cell_Site_Stop_MouseClick(object sender, MouseEventArgs e)
        {
            Switch_Gb_Site_Stop(false, true);
            tb_Groupbox_Cell_Site_Stop.Focus();
        }

        private void btn_Cancel_Click(object sender, EventArgs e)
        {
            if (Check_EmptyValues())
            {
                Application.Exit();
            }
            All_Clear_Data();
            
        }
    }// end form1
}