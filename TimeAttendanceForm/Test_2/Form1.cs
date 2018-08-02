using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace Test_2
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
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
            DialogResult result = openFileDialog1.ShowDialog(); // Show the dialog.
            if (result == DialogResult.OK) // Test result.
            {
                string Pathfile = openFileDialog1.FileName;
                try
                {
                    string text = File.ReadAllText(Pathfile);
                    tb_Target_file.Text = @"...\" + Path.GetFileName(Pathfile);
                    
                }
                catch (IOException)
                {
                }
            }
        }

        private void btn_Preview_Click(object sender, EventArgs e)
        {

        }
    }
}
