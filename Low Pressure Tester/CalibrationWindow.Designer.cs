using OfficeOpenXml;
using System;
using System.IO;
using System.Windows.Forms;

namespace Low_Pressure_Tester
{
    partial class CalibrationWindow
    {
        private void CalibrationWindow_Load(object sender, System.EventArgs e)
        {
            RetestersInitBox.Focus();
        }
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

        // Select next control on enter keypress.

        private void Control_KeyUp(object sender, KeyEventArgs e)
        {
            if ((e.KeyCode == Keys.Enter) || (e.KeyCode == Keys.Return || e.KeyCode == Keys.Tab))
            {
                this.SelectNextControl((System.Windows.Forms.Control)sender, true, true, true, true);
                e.Handled = true;
                e.SuppressKeyPress = true;
            }
        }

        // Run this to save info to the file.
        public static void SaveInfo()
        {
            // Set folder location.
            string path = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            string dirPath = path + "/Check-In Hydrotest Files/";
            DirectoryInfo di = Directory.CreateDirectory(dirPath);
            DateTime now = DateTime.Now;
            int month = now.Month;
            int day = now.Day;
            int year = now.Year;
            // Create file with the date as the file name to match old filing system.
            string date = month + "-" + day + "-" + year;
            string fullPath = dirPath + date + ".xlsx";
            if (File.Exists(fullPath) == false)
            {
                // Initialize Hydrotest Worksheet, add and format header row.
                string appPath = Path.GetDirectoryName(Application.ExecutablePath);
                string templatePath = appPath + "/low pressure sheet.xlsx";
                templatePath = templatePath.Replace("/", @"\");
                System.IO.File.Copy(templatePath, fullPath);
            }
            FileInfo hydroFile = new FileInfo(fullPath);
            using (ExcelPackage hydroPkg = new ExcelPackage(hydroFile))
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                ExcelWorksheet hydroCalibSheet = hydroPkg.Workbook.Worksheets[0];
                int rowNum = 3;
                bool endLoop = false;
                while (endLoop != true)
                {
                    var emptyCell = hydroCalibSheet.Cells["B" + rowNum].Value;
                    if (emptyCell == null)
                    {
                        endLoop = true;
                    }
                    else
                    {
                        rowNum++;
                        endLoop = false;
                    }
                }
                hydroCalibSheet.Cells["B" + rowNum].Value = masterPSI;
                hydroCalibSheet.Cells["C" + rowNum].Value = workingPSI;
                hydroCalibSheet.Cells["E" + rowNum].Value = retesterInit;
                hydroPkg.SaveAs(hydroFile);
                hydroCalibSheet.Calculate();
                // This section is to see if the % deviation is too great to test.
                string deviationString = hydroCalibSheet.Cells["D" + rowNum].Value.ToString();
                deviationString = deviationString.Replace("%", "");
                bool devBool = Decimal.TryParse(deviationString, out decimal deviationDec);
                double deviation = Convert.ToDouble(Math.Round(deviationDec, 2));
                if (deviation > 0.05 || deviation < -0.05)
                {
                    MessageBox.Show("Please verify that testing gauges are accurate.", "ERROR: Deviation percentage greater than 0.5%.", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    hydroCalibSheet.Cells["B" + rowNum].Clear();
                    hydroCalibSheet.Cells["C" + rowNum].Clear();
                    hydroCalibSheet.Cells["E" + rowNum].Clear();
                    hydroPkg.SaveAs(hydroFile);

                }
                LowPressureTester.isCalibrated = true;
            }
        }
        // Run this to clear the boxes and put focus on the first field again.
        private void Control_ReturnFocus(object sender, KeyEventArgs e)
        {
            if ((e.KeyCode == Keys.Enter) || (e.KeyCode == Keys.Return || e.KeyCode == Keys.Tab))
            {
                SaveInfo();
                MasterPressureBox.Clear();
                WorkingPressureBox.Clear();
                MasterPressureBox.Focus();
                e.Handled = true;
                e.SuppressKeyPress = true;
            }
        }
        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.panel1 = new System.Windows.Forms.Panel();
            this.label4 = new System.Windows.Forms.Label();
            this.RetestersInitBox = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.WorkingPressureBox = new System.Windows.Forms.TextBox();
            this.MasterPressureBox = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.panel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.label4);
            this.panel1.Controls.Add(this.RetestersInitBox);
            this.panel1.Controls.Add(this.label3);
            this.panel1.Controls.Add(this.label2);
            this.panel1.Controls.Add(this.WorkingPressureBox);
            this.panel1.Controls.Add(this.MasterPressureBox);
            this.panel1.Controls.Add(this.label1);
            this.panel1.Location = new System.Drawing.Point(13, 13);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(413, 144);
            this.panel1.TabIndex = 0;
            // 
            // label4
            // 
            this.label4.Anchor = System.Windows.Forms.AnchorStyles.Top;
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(165, 49);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(86, 13);
            this.label4.TabIndex = 7;
            this.label4.Text = "Retester\'s Initials";
            // 
            // RetestersInitBox
            // 
            this.RetestersInitBox.Anchor = System.Windows.Forms.AnchorStyles.Top;
            this.RetestersInitBox.Location = new System.Drawing.Point(186, 65);
            this.RetestersInitBox.Name = "RetestersInitBox";
            this.RetestersInitBox.Size = new System.Drawing.Size(42, 20);
            this.RetestersInitBox.TabIndex = 6;
            this.RetestersInitBox.TextChanged += new System.EventHandler(this.RetestersInitBox_TextChanged);
            // 
            // label3
            // 
            this.label3.Anchor = System.Windows.Forms.AnchorStyles.Top;
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(251, 88);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(126, 13);
            this.label3.TabIndex = 5;
            this.label3.Text = "Working Gauge Pressure";
            // 
            // label2
            // 
            this.label2.Anchor = System.Windows.Forms.AnchorStyles.Top;
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(56, 88);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(118, 13);
            this.label2.TabIndex = 4;
            this.label2.Text = "Master Gauge Pressure";
            // 
            // WorkingPressureBox
            // 
            this.WorkingPressureBox.Anchor = System.Windows.Forms.AnchorStyles.Top;
            this.WorkingPressureBox.Location = new System.Drawing.Point(259, 104);
            this.WorkingPressureBox.Name = "WorkingPressureBox";
            this.WorkingPressureBox.Size = new System.Drawing.Size(100, 20);
            this.WorkingPressureBox.TabIndex = 3;
            this.WorkingPressureBox.TextChanged += new System.EventHandler(this.WorkingPressureBox_TextChanged);
            this.WorkingPressureBox.KeyDown += new System.Windows.Forms.KeyEventHandler(this.Control_ReturnFocus);
            // 
            // MasterPressureBox
            // 
            this.MasterPressureBox.Anchor = System.Windows.Forms.AnchorStyles.Top;
            this.MasterPressureBox.Location = new System.Drawing.Point(65, 104);
            this.MasterPressureBox.Name = "MasterPressureBox";
            this.MasterPressureBox.Size = new System.Drawing.Size(100, 20);
            this.MasterPressureBox.TabIndex = 2;
            this.MasterPressureBox.TextChanged += new System.EventHandler(this.MasterPressureBox_TextChanged);
            this.MasterPressureBox.KeyDown += new System.Windows.Forms.KeyEventHandler(this.Control_KeyUp);
            // 
            // label1
            // 
            this.label1.Anchor = System.Windows.Forms.AnchorStyles.Top;
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 20.25F, System.Drawing.FontStyle.Bold);
            this.label1.Location = new System.Drawing.Point(58, 10);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(301, 31);
            this.label1.TabIndex = 1;
            this.label1.Text = "Enter calibration data.";
            // 
            // CalibrationWindow
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(438, 169);
            this.Controls.Add(this.panel1);
            this.Name = "CalibrationWindow";
            this.Text = "Calibration Window";
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox RetestersInitBox;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox WorkingPressureBox;
        private System.Windows.Forms.TextBox MasterPressureBox;
    }
}