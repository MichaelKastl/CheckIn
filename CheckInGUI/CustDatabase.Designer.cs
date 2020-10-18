using DocumentFormat.OpenXml.Spreadsheet;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace CheckInGUI
{

    partial class CustDatabase
    {
        // Designer.
        private System.ComponentModel.IContainer components = null;

        /// Global/Public variables.
        public HashSet<string> custNames = new HashSet<string>();

        // Dispose of resources.
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }
        // Checks for existence of a spreadsheet.
        internal static bool SheetExist(string fullFilePath, string sheetName)
        {
            using (var package = new ExcelPackage(new FileInfo(fullFilePath)))
            {
                return package.Workbook.Worksheets.Any(sheet => sheet.Name == sheetName);
            }
        }
        // Validates file names.
        public string GetSafeFilename(string filename) 
        { 
            return string.Join("_", filename.Split(Path.GetInvalidFileNameChars())); 
        }
        // Makes Excel package, reads Excel file, finds customer data, finds extg data, makes customer database files, and appends extg data to customer database files.
        public void DatabaseBuilder()
        {
            // Tests for if the user has selected a CIDB (Check-In Database) file.
            if (fileChosen == true)
            {
                // Make Excel package to lookup customer names.
                FileInfo excelFile = new FileInfo(file);
                using (ExcelPackage excel = new ExcelPackage(excelFile))
                {
                    ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                    ExcelWorksheet workSheet = excel.Workbook.Worksheets[0];
                    // Iterate through the document.
                    var start = 2;
                    var end = workSheet.Dimension.End;
                    for (int row = start; row <= end.Row; row++)
                    {
                        // Append the Customer Name and Customer Town HashSets with the names and towns from each cell in their respective rows.
                        string custName = workSheet.Cells["O"+row].Text;
                        string custTown = workSheet.Cells["P"+row].Text;
                        string fullName = custName + " - " + custTown;
                        custNames.Add(fullName);
                    };
                    // Write the Database folder.
                    string path = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                    string dirPathParent = path + "/Check-In Database/";
                    DirectoryInfo di = Directory.CreateDirectory(dirPathParent);
                    string dirPath = dirPathParent + "/Customer Database/";
                    DirectoryInfo dir = Directory.CreateDirectory(dirPath);
                    // Iterate through the custNames HashSet, and create a new Excel file in the database for each entry.
                    foreach (string varName in custNames)
                    {
                        // When it's working, write to MegaBase.

                        MegaBaseBuilder.MegaBaseBuilderMethod(workSheet, varName);
                        
                        
                        // Create a file for each customer.
                        string name = GetSafeFilename(varName);
                        string fullPath = dirPath + name + ".xlsx";
                        FileInfo custFile = new FileInfo(fullPath);
                        using (ExcelPackage custPkg = new ExcelPackage(custFile))
                        {
                            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                            // Test for the worksheet existing in the customer file. If not, make one.
                            if (SheetExist(fullPath, "Extinguishers") == false)
                            {
                                custPkg.Workbook.Worksheets.Add("Extinguishers");
                            }
                            // Initialize Customer Worksheet, add and format header row.
                            ExcelWorksheet custSheet = custPkg.Workbook.Worksheets[0];
                            custSheet.Cells["A1"].Value = "Model Num.";
                            custSheet.Cells["B1"].Value = "Serial Num.";
                            custSheet.Cells["A1:B1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            custSheet.Cells["A1:B1"].Style.Font.Size = 14;
                            custSheet.Cells["A1:B1"].Style.Font.Bold = true;
                            custSheet.Cells["A1:B1"].AutoFitColumns();

                            // Start by trimming the name variable to just the Customer Name.
                            int rowStart = workSheet.Dimension.Start.Row;
                            int rowEnd = workSheet.Dimension.End.Row + 1;

                            string cellRange = "A" + rowStart + ":" + "X" + rowEnd;
                            string custName;
                            int index = name.LastIndexOf("-")-1;
                            custName = name.Substring(0, index);

                            // Now, search for the first occurence of that Customer Name in the document, as well as how many times it repeats.
                            var searchCell = from cell in workSheet.Cells[cellRange]
                                                where cell.Value != null && cell.Value.ToString() == custName && cell.Start != null
                                                select cell.Start.Row;
                            var countCell = workSheet.Cells[cellRange].Count(c => c.Text == custName);
                            var searchCell3 = workSheet.Cells[cellRange].FirstOrDefault(c => c.Text == custName);
                            if (custName == null)
                            {
                                searchCell = null;
                            }

                            int? rowNum;
                            int? nameCount;
                            if (searchCell != null)
                            {
                                nameCount = countCell;
                                if (nameCount != null)
                                {
                                    
                                    rowNum = searchCell.FirstOrDefault();
                                }
                                else
                                {
                                    nameCount = 0;
                                    rowNum = 1;
                                }
                            }

                            else
                            {
                                nameCount = 0;
                                rowNum = 1;
                            }
                            int rowCnt = 2;

                            // Loop for as many times as the Customer Name appears.

                            if (nameCount > 0)
                            {
                                while (nameCount > 0)
                                {
                                    // Set modelNum and serialNum to the values of the cells in the rows where Customer Name was found.
                                    string modelNum = workSheet.Cells["A" + rowNum].Value.ToString();
                                    string serialNum = workSheet.Cells["B" + rowNum].Value.ToString();


                                    custSheet.Cells["A" + rowCnt].Value = modelNum;
                                    custSheet.Cells["B" + rowCnt].Value = serialNum;

                                    rowCnt++;



                                    if (nameCount >= 1)
                                    {
                                        string cellRange2 = "A" + rowNum + ":" + "X" + rowEnd;
                                        var searchCell2 = from cell in workSheet.Cells[cellRange2]
                                                          where cell.Value != null && cell.Value.ToString() == custName && cell.Start != null
                                                          select cell.Start.Row;
                                        var searchCell4 = workSheet.Cells[cellRange].FirstOrDefault(c => c.Text == custName);

                                        // Find the next occurence of Customer Name.
                                        if (searchCell2 != null)
                                        {                                           
                                            rowNum = Convert.ToInt32(searchCell2.FirstOrDefault());                                            
                                        }
                                        else
                                        {
                                        }
                                    }
                                    // Set the range to 1 row past the previous occurence of Customer Name.
                                    rowNum++;
                                    // Save the customer file.
                                    custPkg.SaveAs(custFile);

                                    // Decrement the loop counter, and go again.
                                    nameCount--;
                                }
                            }
                            
                        }
                    }
                }
                // Show success dialogbox.
                MessageBox.Show("Customer databases created.", "SUCCESS: Database retrieval complete.",
                   MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                // Error message, in case the user hasn't chosen a database.
                MessageBox.Show("No databas selected, defaulted to the currently active database. If you want to select a different database, please click \"Browse...\" and select a check-in database in order to build a Customer Database.", "INFO: No CIDB selected.",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                string pathDef = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                DateTime nowDef = DateTime.Now;
                string monthDef = DateTime.Now.ToString("MMMM");
                int dayDef = nowDef.Day;
                int yearDef = nowDef.Year;
                string dirPathDef = pathDef + "/Check-In Library/";
                DirectoryInfo di = Directory.CreateDirectory(dirPathDef);
                // Create file with the date as the file name to match old filing system.
                string fullPathDef = dirPathDef + monthDef + " " + yearDef + " Check-In.xlsx";
                file = fullPathDef;
                fileChosen = true;
                DatabaseBuilder();
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
            this.buildDatabaseButton = new System.Windows.Forms.Button();
            this.selectDBFileButton = new System.Windows.Forms.Button();
            this.custDatabaseTitle = new System.Windows.Forms.Label();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.panel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.buildDatabaseButton);
            this.panel1.Controls.Add(this.selectDBFileButton);
            this.panel1.Controls.Add(this.custDatabaseTitle);
            this.panel1.Location = new System.Drawing.Point(12, 12);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(776, 155);
            this.panel1.TabIndex = 0;
            // 
            // buildDatabaseButton
            // 
            this.buildDatabaseButton.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F);
            this.buildDatabaseButton.Location = new System.Drawing.Point(509, 100);
            this.buildDatabaseButton.Name = "buildDatabaseButton";
            this.buildDatabaseButton.Size = new System.Drawing.Size(200, 50);
            this.buildDatabaseButton.TabIndex = 2;
            this.buildDatabaseButton.Text = "Build &Database";
            this.buildDatabaseButton.UseVisualStyleBackColor = true;
            this.buildDatabaseButton.Click += new System.EventHandler(this.buildDatabaseButton_Click);
            // 
            // selectDBFileButton
            // 
            this.selectDBFileButton.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F);
            this.selectDBFileButton.Location = new System.Drawing.Point(113, 100);
            this.selectDBFileButton.Name = "selectDBFileButton";
            this.selectDBFileButton.Size = new System.Drawing.Size(200, 50);
            this.selectDBFileButton.TabIndex = 1;
            this.selectDBFileButton.Text = "&Browse...";
            this.selectDBFileButton.UseVisualStyleBackColor = true;
            this.selectDBFileButton.Click += new System.EventHandler(this.selectDBFileButton_Click);
            // 
            // custDatabaseTitle
            // 
            this.custDatabaseTitle.AutoSize = true;
            this.custDatabaseTitle.Font = new System.Drawing.Font("Microsoft Sans Serif", 20.25F, System.Drawing.FontStyle.Bold);
            this.custDatabaseTitle.Location = new System.Drawing.Point(19, 12);
            this.custDatabaseTitle.Name = "custDatabaseTitle";
            this.custDatabaseTitle.Size = new System.Drawing.Size(736, 31);
            this.custDatabaseTitle.TabIndex = 0;
            this.custDatabaseTitle.Text = "Choose a Monthly Database to Collect Cust. Data From";
            // 
            // CustDatabase
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 176);
            this.Controls.Add(this.panel1);
            this.Name = "CustDatabase";
            this.Text = "Customer Database Builder";
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Label custDatabaseTitle;
        private System.Windows.Forms.Button buildDatabaseButton;
        private System.Windows.Forms.Button selectDBFileButton;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
    }
}