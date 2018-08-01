namespace TimeAttendanceForm
{
    partial class Form_Home
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.Btn_BrowseFile_1 = new System.Windows.Forms.Button();
            this.Pn_file = new System.Windows.Forms.Panel();
            this.textBox3 = new System.Windows.Forms.TextBox();
            this.textBox2 = new System.Windows.Forms.TextBox();
            this.Btn_BrowseFile_2 = new System.Windows.Forms.Button();
            this.Btn_BrowseFile_3 = new System.Windows.Forms.Button();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.lb_TemplateFile = new System.Windows.Forms.Label();
            this.lb_DestFile = new System.Windows.Forms.Label();
            this.lb_TargetFile = new System.Windows.Forms.Label();
            this.panel1 = new System.Windows.Forms.Panel();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.Pn_file.SuspendLayout();
            this.SuspendLayout();
            // 
            // Btn_BrowseFile_1
            // 
            this.Btn_BrowseFile_1.Location = new System.Drawing.Point(792, 13);
            this.Btn_BrowseFile_1.Name = "Btn_BrowseFile_1";
            this.Btn_BrowseFile_1.Size = new System.Drawing.Size(35, 31);
            this.Btn_BrowseFile_1.TabIndex = 0;
            this.Btn_BrowseFile_1.Text = "...";
            this.Btn_BrowseFile_1.UseVisualStyleBackColor = true;
            this.Btn_BrowseFile_1.Click += new System.EventHandler(this.Btn_BrowseFile_1_Click);
            // 
            // Pn_file
            // 
            this.Pn_file.Controls.Add(this.textBox3);
            this.Pn_file.Controls.Add(this.textBox2);
            this.Pn_file.Controls.Add(this.Btn_BrowseFile_2);
            this.Pn_file.Controls.Add(this.Btn_BrowseFile_3);
            this.Pn_file.Controls.Add(this.textBox1);
            this.Pn_file.Controls.Add(this.lb_TemplateFile);
            this.Pn_file.Controls.Add(this.lb_DestFile);
            this.Pn_file.Controls.Add(this.lb_TargetFile);
            this.Pn_file.Controls.Add(this.Btn_BrowseFile_1);
            this.Pn_file.Dock = System.Windows.Forms.DockStyle.Top;
            this.Pn_file.Location = new System.Drawing.Point(0, 0);
            this.Pn_file.Name = "Pn_file";
            this.Pn_file.Size = new System.Drawing.Size(839, 208);
            this.Pn_file.TabIndex = 1;
            // 
            // textBox3
            // 
            this.textBox3.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(222)));
            this.textBox3.Location = new System.Drawing.Point(165, 137);
            this.textBox3.Name = "textBox3";
            this.textBox3.Size = new System.Drawing.Size(607, 20);
            this.textBox3.TabIndex = 8;
            // 
            // textBox2
            // 
            this.textBox2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(222)));
            this.textBox2.Location = new System.Drawing.Point(165, 97);
            this.textBox2.Name = "textBox2";
            this.textBox2.Size = new System.Drawing.Size(607, 20);
            this.textBox2.TabIndex = 7;
            // 
            // Btn_BrowseFile_2
            // 
            this.Btn_BrowseFile_2.Location = new System.Drawing.Point(792, 91);
            this.Btn_BrowseFile_2.Name = "Btn_BrowseFile_2";
            this.Btn_BrowseFile_2.Size = new System.Drawing.Size(35, 31);
            this.Btn_BrowseFile_2.TabIndex = 6;
            this.Btn_BrowseFile_2.Text = "...";
            this.Btn_BrowseFile_2.UseVisualStyleBackColor = true;
            // 
            // Btn_BrowseFile_3
            // 
            this.Btn_BrowseFile_3.Location = new System.Drawing.Point(792, 131);
            this.Btn_BrowseFile_3.Name = "Btn_BrowseFile_3";
            this.Btn_BrowseFile_3.Size = new System.Drawing.Size(35, 31);
            this.Btn_BrowseFile_3.TabIndex = 5;
            this.Btn_BrowseFile_3.Text = "...";
            this.Btn_BrowseFile_3.UseVisualStyleBackColor = true;
            // 
            // textBox1
            // 
            this.textBox1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(222)));
            this.textBox1.Location = new System.Drawing.Point(165, 19);
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(607, 20);
            this.textBox1.TabIndex = 4;
            // 
            // lb_TemplateFile
            // 
            this.lb_TemplateFile.AutoSize = true;
            this.lb_TemplateFile.Font = new System.Drawing.Font("Microsoft Sans Serif", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(222)));
            this.lb_TemplateFile.Location = new System.Drawing.Point(12, 128);
            this.lb_TemplateFile.Name = "lb_TemplateFile";
            this.lb_TemplateFile.Size = new System.Drawing.Size(135, 25);
            this.lb_TemplateFile.TabIndex = 3;
            this.lb_TemplateFile.Text = "Template file";
            // 
            // lb_DestFile
            // 
            this.lb_DestFile.AutoSize = true;
            this.lb_DestFile.Font = new System.Drawing.Font("Microsoft Sans Serif", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(222)));
            this.lb_DestFile.Location = new System.Drawing.Point(12, 88);
            this.lb_DestFile.Name = "lb_DestFile";
            this.lb_DestFile.Size = new System.Drawing.Size(90, 25);
            this.lb_DestFile.TabIndex = 2;
            this.lb_DestFile.Text = "Dest file";
            // 
            // lb_TargetFile
            // 
            this.lb_TargetFile.AutoSize = true;
            this.lb_TargetFile.Font = new System.Drawing.Font("Microsoft Sans Serif", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(222)));
            this.lb_TargetFile.Location = new System.Drawing.Point(12, 13);
            this.lb_TargetFile.Name = "lb_TargetFile";
            this.lb_TargetFile.Size = new System.Drawing.Size(108, 25);
            this.lb_TargetFile.TabIndex = 1;
            this.lb_TargetFile.Text = "Target file";
            // 
            // panel1
            // 
            this.panel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel1.Location = new System.Drawing.Point(0, 208);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(839, 353);
            this.panel1.TabIndex = 2;
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // Form_Home
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(839, 561);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.Pn_file);
            this.Name = "Form_Home";
            this.Text = "C.S.I. Group - Time Attendance";
            this.Pn_file.ResumeLayout(false);
            this.Pn_file.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button Btn_BrowseFile_1;
        private System.Windows.Forms.Panel Pn_file;
        private System.Windows.Forms.TextBox textBox3;
        private System.Windows.Forms.TextBox textBox2;
        private System.Windows.Forms.Button Btn_BrowseFile_2;
        private System.Windows.Forms.Button Btn_BrowseFile_3;
        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.Label lb_TemplateFile;
        private System.Windows.Forms.Label lb_DestFile;
        private System.Windows.Forms.Label lb_TargetFile;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
    }
}

