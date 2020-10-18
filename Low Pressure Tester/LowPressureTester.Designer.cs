using OfficeOpenXml;
using System;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using System.Windows.Media;

namespace Low_Pressure_Tester
{
    partial class LowPressureTester
    {
        public static bool isCalibrated;
        public static string custNameLookup;
        public static string custTownLookup;
        // Designer
        private System.ComponentModel.IContainer components = null;

        // Clear resources.
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        // This method tests for the existence of a Hydrotest calibration sheet for the day. 

        public static void CalibrationData()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
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
            if (File.Exists(fullPath) == true)
            { 
                FileInfo hydroFile = new FileInfo(fullPath);
                using (ExcelPackage hydroPkg = new ExcelPackage(hydroFile))
                {
                    ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                    ExcelWorksheet hydroCalibSheet = hydroPkg.Workbook.Worksheets[0];
                    if (hydroCalibSheet.Cells["B3"].Value != null)
                    {
                        isCalibrated = true;
                    }
                    else
                    {
                        isCalibrated = false;
                    }
                }
            } 
            else
            {
                isCalibrated = false;
            }
        }

        // This method clears the data in the text boxes and returns the focus to the first field.
        private void Control_ReturnFocus(object sender, KeyEventArgs e)
        {
            if ((e.KeyCode == Keys.Enter) || (e.KeyCode == Keys.Return || e.KeyCode == Keys.Tab))
            {
                CalibrationData();
                if (isCalibrated == true)
                {
                    custTownBox.Clear();
                    custNameBox.Clear();
                    CustomerLookup(serialNum);
                    if (custNameLookup != null)
                    {
                        customer = custNameLookup;
                        custNameBox.Text = customer;
                        custTown = custTownLookup;
                        custTownBox.Text = custTown;
                        HydroFileWriter(customer, custTown, serialNum, modelNum);
                    } else if (manualCustDataCheck.Checked == true)
                    {
                        customer = custNameBox.Text;
                        custTown = custTownBox.Text;
                        serialNum = serialNumBox.Text;
                        modelNum = modelNumBox.Text;
                        HydroFileWriter(customer, custTown, serialNum, modelNum);
                    }
                    modelNumBox.Clear();
                    serialNumBox.Clear();
                    modelNumBox.Focus();
                }
                else
                {
                    MessageBox.Show("Please calibrate the tester before entering extinguisher data.", "ERROR: No calibration data found.", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
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

        // Checks for existence of a spreadsheet.
        internal static bool SheetExist(string fullFilePath, string sheetName)
        {
            using (var package = new ExcelPackage(new FileInfo(fullFilePath)))
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                return package.Workbook.Worksheets.Any(sheet => sheet.Name == sheetName);
            }
        }

        // Method to read through the SerialNumDatabase for the customer that the extinguisher belongs to.
        public static void CustomerLookup(string serialLookup)
        {
            // Create Mega Database path.
            string docPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            string megaDirPath = docPath + "/Check-In Database/Serial Database/";
            DirectoryInfo dir = Directory.CreateDirectory(megaDirPath);
            string megaPath = megaDirPath + "SerialNumDatabase.xlsx";
            // Create ExcelFile for Mega Database.
            FileInfo megaFile = new FileInfo(megaPath);
            // Open Mega Database package.
            using (ExcelPackage megaPkg = new ExcelPackage(megaFile))
            {
                // License for EPPlus.
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                // Initialize worksheet, and set worksheet dimensions.
                ExcelWorksheet megaWorksheet = megaPkg.Workbook.Worksheets[0];
                int rowStart = 2;
                int rowEnd = megaWorksheet.Dimension.End.Row;
                string cellRange = "A" + rowStart + ":D" + rowEnd;

                // Look for the row the serial number is in.
                var searchCell = from cell in megaWorksheet.Cells[cellRange]
                                 where cell.Value != null && cell.Value.ToString() == serialNum && cell.Start != null
                                 select cell.Start.Row;
                if (searchCell != null)
                {
                    int? serialLocation = searchCell.FirstOrDefault();
                    // Store the values of the Customer Name and Customer Town cells in the serial number's row.
                    if (serialLocation != null && serialLocation != 0)
                    {
                        if (megaWorksheet.Cells["A" + serialLocation].Value.ToString() != null)
                        {
                            custNameLookup = megaWorksheet.Cells["A" + serialLocation].Value.ToString();
                            custTownLookup = megaWorksheet.Cells["B" + serialLocation].Value.ToString();
                        }
                    }
                    else
                    {
                        custNameLookup = null;
                        custTownLookup = null;
                        MessageBox.Show("No customer data found matching this serial number. Please double-check your serial number, or enter customer information manually.", "ERROR: No Cust. Match Found", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }              
            }
        }

        // Method to create HT form and write data to it.
        public static void HydroFileWriter(string customer, string custTown, string serialNum, string modelNum)
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
                templatePath = templatePath.Replace(@"\", "/");
                Console.WriteLine(templatePath);
                Console.WriteLine(fullPath);
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
                    var emptyCell = hydroTestingSheet.Cells["B" + rowNum].Value;
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
                hydroTestingSheet.Cells["B" + rowNum].Value = serialNum;
                hydroTestingSheet.Cells["C" + rowNum].Value = customer;
                hydroTestingSheet.Cells["D" + rowNum].Value = custTown;
                hydroTestingSheet.Cells["F" + rowNum].Value = modelNum;
                hydroPkg.SaveAs(hydroFile);
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
            this.calibrationModeButton = new System.Windows.Forms.Button();
            this.testModeButton = new System.Windows.Forms.Button();
            this.label5 = new System.Windows.Forms.Label();
            this.manualCustDataCheck = new System.Windows.Forms.CheckBox();
            this.label4 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.custTownBox = new System.Windows.Forms.TextBox();
            this.custNameBox = new System.Windows.Forms.TextBox();
            this.serialNumBox = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.modelNumBox = new System.Windows.Forms.TextBox();
            this.panel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.calibrationModeButton);
            this.panel1.Controls.Add(this.testModeButton);
            this.panel1.Controls.Add(this.label5);
            this.panel1.Controls.Add(this.manualCustDataCheck);
            this.panel1.Controls.Add(this.label4);
            this.panel1.Controls.Add(this.label3);
            this.panel1.Controls.Add(this.label2);
            this.panel1.Controls.Add(this.custTownBox);
            this.panel1.Controls.Add(this.custNameBox);
            this.panel1.Controls.Add(this.serialNumBox);
            this.panel1.Controls.Add(this.label1);
            this.panel1.Controls.Add(this.modelNumBox);
            this.panel1.Location = new System.Drawing.Point(12, 12);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(776, 170);
            this.panel1.TabIndex = 0;
            // 
            // calibrationModeButton
            // 
            this.calibrationModeButton.Location = new System.Drawing.Point(238, 130);
            this.calibrationModeButton.Name = "calibrationModeButton";
            this.calibrationModeButton.Size = new System.Drawing.Size(105, 40);
            this.calibrationModeButton.TabIndex = 11;
            this.calibrationModeButton.TabStop = false;
            this.calibrationModeButton.Text = "Enter &Calibration Mode";
            this.calibrationModeButton.UseVisualStyleBackColor = true;
            this.calibrationModeButton.Click += new System.EventHandler(this.calibrationModeButton_Click);
            // 
            // testModeButton
            // 
            this.testModeButton.Location = new System.Drawing.Point(435, 130);
            this.testModeButton.Name = "testModeButton";
            this.testModeButton.Size = new System.Drawing.Size(105, 40);
            this.testModeButton.TabIndex = 10;
            this.testModeButton.TabStop = false;
            this.testModeButton.Text = "Enter &Test Mode";
            this.testModeButton.UseVisualStyleBackColor = true;
            this.testModeButton.Click += new System.EventHandler(this.testModeButton_Click);
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Font = new System.Drawing.Font("Microsoft Sans Serif", 20.25F, System.Drawing.FontStyle.Bold);
            this.label5.Location = new System.Drawing.Point(189, 9);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(386, 31);
            this.label5.TabIndex = 9;
            this.label5.Text = "Low Pressure Test Recorder";
            // 
            // manualCustDataCheck
            // 
            this.manualCustDataCheck.Location = new System.Drawing.Point(315, 100);
            this.manualCustDataCheck.Name = "manualCustDataCheck";
            this.manualCustDataCheck.Size = new System.Drawing.Size(155, 17);
            this.manualCustDataCheck.TabIndex = 8;
            this.manualCustDataCheck.TabStop = false;
            this.manualCustDataCheck.Text = "Manually Enter Cust. Data";
            this.manualCustDataCheck.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            this.manualCustDataCheck.UseVisualStyleBackColor = true;
            this.manualCustDataCheck.CheckedChanged += new System.EventHandler(this.manualCustDataCheck_CheckedChanged);
            // 
            // label4
            // 
            this.label4.Location = new System.Drawing.Point(645, 84);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(100, 13);
            this.label4.TabIndex = 7;
            this.label4.Text = "Town";
            this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            this.label4.Click += new System.EventHandler(this.label4_Click);
            // 
            // label3
            // 
            this.label3.Location = new System.Drawing.Point(440, 84);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(100, 13);
            this.label3.TabIndex = 6;
            this.label3.Text = "Customer";
            this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            this.label3.Click += new System.EventHandler(this.label3_Click);
            // 
            // label2
            // 
            this.label2.Location = new System.Drawing.Point(235, 85);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(100, 13);
            this.label2.TabIndex = 5;
            this.label2.Text = "Serial Num";
            this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            this.label2.Click += new System.EventHandler(this.label2_Click);
            // 
            // custTownBox
            // 
            this.custTownBox.Location = new System.Drawing.Point(645, 63);
            this.custTownBox.Name = "custTownBox";
            this.custTownBox.Size = new System.Drawing.Size(100, 20);
            this.custTownBox.TabIndex = 3;
            this.custTownBox.TextChanged += new System.EventHandler(this.custTownBox_TextChanged);
            this.custTownBox.Enter += new System.EventHandler(this.custTownBox_Enter);
            this.custTownBox.KeyDown += new System.Windows.Forms.KeyEventHandler(this.Control_ReturnFocus);
            // 
            // custNameBox
            // 
            this.custNameBox.Location = new System.Drawing.Point(440, 63);
            this.custNameBox.Name = "custNameBox";
            this.custNameBox.Size = new System.Drawing.Size(100, 20);
            this.custNameBox.TabIndex = 2;
            this.custNameBox.TextChanged += new System.EventHandler(this.custNameBox_TextChanged);
            this.custNameBox.Enter += new System.EventHandler(this.custNameBox_Enter);
            this.custNameBox.KeyDown += new System.Windows.Forms.KeyEventHandler(this.Control_KeyUp);
            // 
            // serialNumBox
            // 
            this.serialNumBox.Location = new System.Drawing.Point(235, 63);
            this.serialNumBox.Name = "serialNumBox";
            this.serialNumBox.Size = new System.Drawing.Size(100, 20);
            this.serialNumBox.TabIndex = 1;
            this.serialNumBox.TextChanged += new System.EventHandler(this.serialNumBox_TextChanged);
            this.serialNumBox.Enter += new System.EventHandler(this.serialNumBox_Enter);
            this.serialNumBox.KeyDown += new System.Windows.Forms.KeyEventHandler(this.Control_ReturnFocus);
            // 
            // label1
            // 
            this.label1.Location = new System.Drawing.Point(30, 86);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(100, 13);
            this.label1.TabIndex = 1;
            this.label1.Text = "Model Num.";
            this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            this.label1.Click += new System.EventHandler(this.label1_Click);
            // 
            // modelNumBox
            // 
            this.modelNumBox.Location = new System.Drawing.Point(30, 63);
            this.modelNumBox.Name = "modelNumBox";
            this.modelNumBox.Size = new System.Drawing.Size(100, 20);
            this.modelNumBox.TabIndex = 0;
            this.modelNumBox.TextChanged += new System.EventHandler(this.modelNumBox_TextChanged);
            this.modelNumBox.Enter += new System.EventHandler(this.modelNumBox_Enter);
            this.modelNumBox.KeyDown += new System.Windows.Forms.KeyEventHandler(this.Control_KeyUp);
            // 
            // LowPressureTester
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 186);
            this.Controls.Add(this.panel1);
            this.Name = "LowPressureTester";
            this.Text = "Low Pressure Test Recorder";
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.CheckBox manualCustDataCheck;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox custTownBox;
        private System.Windows.Forms.TextBox custNameBox;
        private System.Windows.Forms.TextBox serialNumBox;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox modelNumBox;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Button testModeButton;
        private Button calibrationModeButton;
    }
}

