using System;
using System.IO;
using System.Windows.Forms;

namespace TimeAttendanceForm
{
    public partial class Form_Home : Form
    {
        public Form_Home()
        {
            InitializeComponent();
        }

        private void Btn_BrowseFile_1_Click(object sender, EventArgs e)
        {
            DialogResult result = openFileDialog1.ShowDialog(); // Show the dialog.
            if (result == DialogResult.OK) // Test result.
            {
                string file = openFileDialog1.FileName;
                try
                {
                    string text = File.ReadAllText(file);
                    textBox1.Text = file;
                }
                catch (IOException)
                {
                }
            }
        }
    }
}
