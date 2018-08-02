using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

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
    }
}
