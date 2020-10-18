using OfficeOpenXml;
using System;
using System.IO;
using System.Windows.Forms;

namespace Low_Pressure_Tester
{
    partial class TestingWindow
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;
        public static bool dataExists;

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

        // This method tests for if any extg. data exists.
        public static void DataExists(string range)
        {
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
            FileInfo hydroFile = new FileInfo(fullPath);
            using (ExcelPackage hydroPkg = new ExcelPackage(hydroFile))
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                ExcelWorksheet hydroTestingSheet = hydroPkg.Workbook.Worksheets[1];
                if (hydroTestingSheet.Cells[range].Value == null)
                {
                    dataExists = false;
                }
                else
                {
                    dataExists = true;
                }
            }
        }

        // This is to return focus to the beginning field.
        private void Control_ReturnFocus(object sender, KeyEventArgs e)
        {
            if ((e.KeyCode == Keys.Enter) || (e.KeyCode == Keys.Return || e.KeyCode == Keys.Tab))
            {
                SaveData();
                TestPSIBox.Clear();
                VisualBox.Clear();
                DispCodeBox.Clear();
                ManufacturerBox.Clear();
                DOTSpecBox.Clear();
                TestPSIBox.Focus();
                e.Handled = true;
                e.SuppressKeyPress = true;
            }
        }
        
        // This is the same method as above, but requiring no KeyEventArgs.
        private void ReturnFocus()
        {
            SaveData();
            TestPSIBox.Clear();
            VisualBox.Clear();
            DispCodeBox.Clear();
            ManufacturerBox.Clear();
            DOTSpecBox.Clear();
            TestPSIBox.Focus();
        }

        // Select next control on enter keypress, unless RetesterInit has already been set, and is the next control. In which case, we'll return focus.

        private void Control_KeyUp(object sender, KeyEventArgs e)
        {
            if ((e.KeyCode == Keys.Enter) || (e.KeyCode == Keys.Return || e.KeyCode == Keys.Tab))
            {
                bool retesterBox = DispCodeBox.Focused;
                string retesterBoxInit = RetesterInitBox.Text;
                if (retesterBox == true && retesterBoxInit != null)
                {
                    ReturnFocus();
                }
                else
                {
                    this.SelectNextControl((System.Windows.Forms.Control)sender, true, true, true, true);
                    e.Handled = true;
                    e.SuppressKeyPress = true;
                }
            }
        }

        // Run this to save the data.
        private void SaveData()
        {
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
                    ExcelWorksheet hydroTestingSheet = hydroPkg.Workbook.Worksheets[1];
                    int rowNum = 3;
                    bool endLoop = false;
                    while (endLoop != true)
                    {
                        var emptyCell = hydroTestingSheet.Cells["H" + rowNum].Value;
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
                    string range = "B" + rowNum;
                    DataExists(range);
                    if (dataExists == true)
                    {
                        hydroTestingSheet.Cells["E" + rowNum].Value = manufacturer;
                        hydroTestingSheet.Cells["G" + rowNum].Value = DOTSpec;
                        hydroTestingSheet.Cells["H" + rowNum].Value = testPSI;
                        hydroTestingSheet.Cells["I" + rowNum].Value = visual;
                        hydroTestingSheet.Cells["J" + rowNum].Value = dispCode;
                        hydroTestingSheet.Cells["K" + rowNum].Value = retesterInit;
                        hydroPkg.SaveAs(hydroFile);
                    }
                    else
                    {
                        hydroTestingSheet.Cells["E" + rowNum].Clear();
                        hydroTestingSheet.Cells["G" + rowNum].Clear();
                        hydroTestingSheet.Cells["H" + rowNum].Clear();
                        hydroTestingSheet.Cells["I" + rowNum].Clear();
                        hydroTestingSheet.Cells["J" + rowNum].Clear();
                        hydroTestingSheet.Cells["K" + rowNum].Clear();
                        hydroPkg.SaveAs(hydroFile);
                        MessageBox.Show("Testing data has surpassed extinguisher data, or no extinguisher data was ever entered.", "ERROR: Data not found.", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                }
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
            this.EnableManuAndDOTCheck = new System.Windows.Forms.CheckBox();
            this.label6 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.DOTSpecBox = new System.Windows.Forms.TextBox();
            this.ManufacturerBox = new System.Windows.Forms.TextBox();
            this.RetesterInitBox = new System.Windows.Forms.TextBox();
            this.DispCodeBox = new System.Windows.Forms.TextBox();
            this.VisualBox = new System.Windows.Forms.TextBox();
            this.TestPSIBox = new System.Windows.Forms.TextBox();
            this.panel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.EnableManuAndDOTCheck);
            this.panel1.Controls.Add(this.label6);
            this.panel1.Controls.Add(this.label5);
            this.panel1.Controls.Add(this.label4);
            this.panel1.Controls.Add(this.label3);
            this.panel1.Controls.Add(this.label2);
            this.panel1.Controls.Add(this.label1);
            this.panel1.Controls.Add(this.DOTSpecBox);
            this.panel1.Controls.Add(this.ManufacturerBox);
            this.panel1.Controls.Add(this.RetesterInitBox);
            this.panel1.Controls.Add(this.DispCodeBox);
            this.panel1.Controls.Add(this.VisualBox);
            this.panel1.Controls.Add(this.TestPSIBox);
            this.panel1.Location = new System.Drawing.Point(13, 13);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(621, 147);
            this.panel1.TabIndex = 0;
            // 
            // EnableManuAndDOTCheck
            // 
            this.EnableManuAndDOTCheck.AutoSize = true;
            this.EnableManuAndDOTCheck.Location = new System.Drawing.Point(243, 126);
            this.EnableManuAndDOTCheck.Name = "EnableManuAndDOTCheck";
            this.EnableManuAndDOTCheck.Size = new System.Drawing.Size(151, 17);
            this.EnableManuAndDOTCheck.TabIndex = 12;
            this.EnableManuAndDOTCheck.Text = "Enable Bottom Text Boxes";
            this.EnableManuAndDOTCheck.UseVisualStyleBackColor = true;
            this.EnableManuAndDOTCheck.CheckedChanged += new System.EventHandler(this.EnableManuAndDOTCheck_CheckedChanged);
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(365, 81);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(61, 13);
            this.label6.TabIndex = 11;
            this.label6.Text = "DOT Spec.";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(194, 81);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(79, 13);
            this.label5.TabIndex = 10;
            this.label5.Text = "Mfg. ID Symbol";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(527, 24);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(47, 13);
            this.label4.TabIndex = 9;
            this.label4.Text = "Retester";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(353, 21);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(86, 13);
            this.label3.TabIndex = 8;
            this.label3.Text = "Disposition Code";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(171, 21);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(124, 13);
            this.label2.TabIndex = 7;
            this.label2.Text = "Visual Inspection (P or F)";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(49, 21);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(72, 13);
            this.label1.TabIndex = 6;
            this.label1.Text = "Test Pressure";
            // 
            // DOTSpecBox
            // 
            this.DOTSpecBox.Location = new System.Drawing.Point(346, 100);
            this.DOTSpecBox.Name = "DOTSpecBox";
            this.DOTSpecBox.Size = new System.Drawing.Size(100, 20);
            this.DOTSpecBox.TabIndex = 5;
            this.DOTSpecBox.TextChanged += new System.EventHandler(this.DOTSpecBox_TextChanged);
            this.DOTSpecBox.KeyDown += new System.Windows.Forms.KeyEventHandler(this.Control_KeyUp);
            // 
            // ManufacturerBox
            // 
            this.ManufacturerBox.Location = new System.Drawing.Point(185, 100);
            this.ManufacturerBox.Name = "ManufacturerBox";
            this.ManufacturerBox.Size = new System.Drawing.Size(100, 20);
            this.ManufacturerBox.TabIndex = 4;
            this.ManufacturerBox.TextChanged += new System.EventHandler(this.ManufacturerBox_TextChanged);
            this.ManufacturerBox.KeyDown += new System.Windows.Forms.KeyEventHandler(this.Control_KeyUp);
            // 
            // RetesterInitBox
            // 
            this.RetesterInitBox.Location = new System.Drawing.Point(502, 40);
            this.RetesterInitBox.Name = "RetesterInitBox";
            this.RetesterInitBox.Size = new System.Drawing.Size(100, 20);
            this.RetesterInitBox.TabIndex = 3;
            this.RetesterInitBox.TextChanged += new System.EventHandler(this.RetesterInitBox_TextChanged);
            this.RetesterInitBox.KeyDown += new System.Windows.Forms.KeyEventHandler(this.Control_ReturnFocus);
            // 
            // DispCodeBox
            // 
            this.DispCodeBox.Location = new System.Drawing.Point(346, 40);
            this.DispCodeBox.Name = "DispCodeBox";
            this.DispCodeBox.Size = new System.Drawing.Size(100, 20);
            this.DispCodeBox.TabIndex = 2;
            this.DispCodeBox.TextChanged += new System.EventHandler(this.DispCodeBox_TextChanged);
            this.DispCodeBox.KeyDown += new System.Windows.Forms.KeyEventHandler(this.Control_KeyUp);
            // 
            // VisualBox
            // 
            this.VisualBox.Location = new System.Drawing.Point(185, 40);
            this.VisualBox.Name = "VisualBox";
            this.VisualBox.Size = new System.Drawing.Size(100, 20);
            this.VisualBox.TabIndex = 1;
            this.VisualBox.TextChanged += new System.EventHandler(this.VisualBox_TextChanged);
            this.VisualBox.KeyDown += new System.Windows.Forms.KeyEventHandler(this.Control_ReturnFocus);
            // 
            // TestPSIBox
            // 
            this.TestPSIBox.Location = new System.Drawing.Point(32, 40);
            this.TestPSIBox.Name = "TestPSIBox";
            this.TestPSIBox.Size = new System.Drawing.Size(100, 20);
            this.TestPSIBox.TabIndex = 0;
            this.TestPSIBox.TextChanged += new System.EventHandler(this.TestPSIBox_TextChanged);
            this.TestPSIBox.KeyDown += new System.Windows.Forms.KeyEventHandler(this.Control_KeyUp);
            // 
            // TestingWindow
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(646, 172);
            this.Controls.Add(this.panel1);
            this.Name = "TestingWindow";
            this.Text = "Testing Window";
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.TextBox DOTSpecBox;
        private System.Windows.Forms.TextBox ManufacturerBox;
        private System.Windows.Forms.TextBox RetesterInitBox;
        private System.Windows.Forms.TextBox DispCodeBox;
        private System.Windows.Forms.TextBox VisualBox;
        private System.Windows.Forms.TextBox TestPSIBox;
        private System.Windows.Forms.CheckBox EnableManuAndDOTCheck;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label1;
    }
}