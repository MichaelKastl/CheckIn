using System;
using System.Windows.Forms;
using System.Linq;
using System.IO;
using OfficeOpenXml;
using System.Collections.Generic;
using OfficeOpenXml.Style;
using System.Xml.Serialization;
using System.Data;
using System.Media;
using DocumentFormat.OpenXml.Spreadsheet;

namespace CheckInGUI
{
    /* General Future Plans
     * 1. Change the template .xmls to .csv rows. No reason to have two data sources.
     * 2. Locate the areas where I used throwaway variables. Make one Extinguisher class for all extinguisher parts and data, and get rid of the throwaway or single-use variables.
     * 3. Fix my bizarre spacing.
     * 4. Methods like control_KeyUp can be generalized. Delete the other methods with the same function from other classes, and just instantiate an object of a single class with those methods.
     */

    partial class CheckIn
    {
        public static string globalExtraParts;
        public static string globalExtraLabor;
        public static bool globalExPartsBool = false;
        public static bool globalExLaborBool = false;

        public static string valveStem;
        public static string oRing1;
        public static string oRing2;
        public static string size;
        public static string chemical;
        public static int htYear;
        public static int sixYear;
        public static bool rbCollar;
        public static bool rbCT;
        public static string extraParts;
        public static string extraLabor;
        public static string rbModel;
        public static string finalWITB;
        public static int currentYear;
        public static int rbYear;
        public static bool rbPSI;
        public static string workNeeded;
        public static string rechRecon;
        public static string ctLabel;
        public static string tagSealCollar;
        public static bool validModel;
        public static bool continueModelLoop;
        public static bool isBlank;
        public static bool showBox;
        public static string headerRange;
        public static string lastModelNumStr;
        DataTable table = new DataTable();

        // Opens an ExcelPackage for the month's sheet for editing.
        public void OpenDoc()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            string month = DateTime.Now.ToString("MMMM");
            int cYear = DateTime.Now.Year;
            string path = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            string dirPath = path + "/Check-In Library/";
            DirectoryInfo di = Directory.CreateDirectory(dirPath);
            string fullPath = dirPath + month + " " + cYear + " Check-In.xlsx";
            FileInfo excelFile2 = new FileInfo(fullPath);
            if (File.Exists(fullPath) == true)
            {
                System.Diagnostics.Process.Start(fullPath);
            }
            else
            {
                SystemSounds.Exclamation.Play();
                MessageBox.Show("Please enter extinguisher data and submit the record to generate this month's database.",
                    "ERROR: Database for this month doesn't exist yet.", MessageBoxButtons.OK, MessageBoxIcon.Warning);

            }
        }

        // Forces cells to calculate and update without opening the spreadsheet.
        public void CalculateSheet()
        {
            string month = DateTime.Now.ToString("MMMM");
            int cYear = DateTime.Now.Year;
            string path = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            string dirPath = path + "/Check-In Library/";
            DirectoryInfo di = Directory.CreateDirectory(dirPath);
            string fullPath = dirPath + month + " " + cYear + " Check-In.xlsx";
            FileInfo excelFile = new FileInfo(fullPath);
            if (File.Exists(fullPath) == true)
            {
                using (ExcelPackage excel = new ExcelPackage(excelFile))
                {
                    ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                    ExcelWorksheet worksheet = excel.Workbook.Worksheets[0];
                    // Output from the logger will be written to the following file
                    var logfile = new FileInfo(dirPath + "logfile.txt");
                    // Attach the logger before the calculation is performed.
                    excel.Workbook.FormulaParserManager.AttachLogger(logfile);
                    // Calculate - can also be executed on sheet- or range level.
                    excel.Workbook.Worksheets[0].Cells["O1:V300"].Calculate();
                    excel.SaveAs(excelFile);
                    ExcelReadoutRefresh();
                    // The following method removes any logger attached to the workbook.
                    excel.Workbook.FormulaParserManager.DetachLogger();
                }
            }
            else
            {
                SystemSounds.Exclamation.Play();
                MessageBox.Show("Please enter extinguisher data and submit the record to generate this month's database.",
                    "ERROR: Database for this month doesn't exist yet.", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            MessageBox.Show("Worksheet formulas run, and values updated to reflect results.",
                "SUCCESS: Calculation Complete.", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        // Forces a refresh of the live-view of the Excel document. Keeps the dataview up-to-date.
        public void ExcelReadoutRefresh()
        {
            string month = DateTime.Now.ToString("MMMM");
            int cYear = DateTime.Now.Year;
            string path = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            string dirPath = path + "/Check-In Library/";
            DirectoryInfo di = Directory.CreateDirectory(dirPath);
            string fullPath = dirPath + month + " " + cYear + " Check-In.xlsx";
            FileInfo excelFile2 = new FileInfo(fullPath);
            if (File.Exists(fullPath))
            {
                using (ExcelPackage excel2 = new ExcelPackage(excelFile2))
                {
                    ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                    DataTable dt = ExcelPackageToDataTable(excel2);
                    dataGridView1.DataSource = dt;
                    dataGridView1.ClearSelection();

                    int nRowIndex = dataGridView1.Rows.Count - 1;

                    dataGridView1.Rows[nRowIndex].Selected = true;
                    dataGridView1.FirstDisplayedScrollingRowIndex = nRowIndex;
                }
            }
        }

        // Translates the open ExcelPackage to the dataviewe for live viewing.
        public static DataTable ExcelPackageToDataTable(ExcelPackage excelPackage)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            DataTable dt = new DataTable();
            ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets[0];
            if (worksheet.Dimension == null)
            {
                return dt;
            }
            List<string> columnNames = new List<string>();
            int currentColumn = 1;
            foreach (var cell in worksheet.Cells[1, 1, 1, worksheet.Dimension.End.Column])
            {
                string columnName = cell.Text.Trim();
                if (cell.Start.Column != currentColumn)
                {
                    columnNames.Add("Header_" + currentColumn);
                    dt.Columns.Add("Header_" + currentColumn);
                    currentColumn++;
                }
                columnNames.Add(columnName);
                int occurrences = columnNames.Count(x => x.Equals(columnName));
                if (occurrences > 1)
                {
                    columnName = columnName + "_" + occurrences;
                }
                dt.Columns.Add(columnName);
                currentColumn++;
            }
            for (int i = 2; i <= worksheet.Dimension.End.Row; i++)
            {
                var row = worksheet.Cells[i, 1, i, worksheet.Dimension.End.Column];
                DataRow newRow = dt.NewRow();
                foreach (var cell in row)
                {
                    newRow[cell.Start.Column - 1] = cell.Text;
                }
                dt.Rows.Add(newRow);
            }
            return dt;
        }

        // Checks for a sheet's existence.
        internal static bool SheetExist(string fullFilePath, string sheetName)
        {
            using (var package = new ExcelPackage(new FileInfo(fullFilePath)))
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                return package.Workbook.Worksheets.Any(sheet => sheet.Name == sheetName);
            }
        }

        // Brings focus to the first text box, as well as highlighting whatever data has been inputted into that box.
        public void BeginningFocus()
        {
            modelNum.Select();
            modelNum.Focus();
        }

        /* This method loads the existing parts database, and checks the database for the correct parts for that model number.
        * If the database does not contain those parts, it checks the custom template XMLs for the right parts for the model number.
        * If neither of those contain the right parts, then the method prompts the user to create a new template for this model number, to ensure future matches.
        *
        * Future plans: Rewrite this so that new templates are added to the .csv instead. No reason to have two different storage methods, and an .xml for each template will cause unnecessary bloat.
        */
        public static void RecordBuilder(string methodModel, string methodSerial, int methodYear, string methodWITB, bool methodPSI)
        {
            rbModel = methodModel.ToUpper();
            currentYear = System.DateTime.Now.Year;
            rbYear = methodYear;
            finalWITB = methodWITB;
            rbPSI = methodPSI;
            string blank = "";
            int cnt = 10;
            List<string> yearSuffixes = new List<string>();
            while (cnt < 100)
            {
                string x = "-" + cnt;
                yearSuffixes.Add(x);
            }
            yearSuffixes.Add("-00");
            cnt = 1;
            while (cnt <= 20)
            {
                string x = "";
                if (cnt <= 9)
                {
                    x = "-0" + cnt;
                }
                else
                {
                    x = "-" + cnt;
                }
                yearSuffixes.Add(x);
            }           

            string stringIndex1 = methodModel.Substring(0, 1);

            if ((stringIndex1 == "A") || (stringIndex1 == "B") || (stringIndex1 == "X")) { methodModel = methodModel.Replace(stringIndex1, blank); }

            if (stringIndex1 != "X")
            {
                string stringIndex2 = methodModel.Substring(0, 1);

                if ((stringIndex2 == "A") || (stringIndex2 == "B")) { methodModel = methodModel.Replace(stringIndex2, blank); }
            }
            int startIndex = methodModel.Length - 3;
            string suffix = methodModel.Substring(startIndex, 3);
            foreach (string x in yearSuffixes)
            {
                if (suffix == x) { methodModel = methodModel.Replace(suffix, blank); }
            }
            lastModelNumStr = methodModel;
            Dictionary<string, CheckInData> database = new Dictionary<string, CheckInData>();
            using (var reader = new StreamReader(@"C:\Users\mkast\source\repos\CheckInGUI\CheckInGUI\extgData.csv"))
            {
                while (!reader.EndOfStream)
                {
                    CheckInData cid = new CheckInData();
                    var line = reader.ReadLine();
                    var values = line.Split(',');
                    if (values[0] != null)
                    {
                        cid.modelNum = values[0];
                        cid.valveStem = values[1];
                        cid.oRing1 = values[2];
                        cid.size = values[3];
                        cid.chemical = values[4];
                        cid.htYear = Convert.ToInt32(values[5]);
                        cid.sixYear = Convert.ToInt32(values[6]);
                        cid.rbCT = Convert.ToBoolean(values[7]);
                        cid.extraParts = values[8];
                        cid.extraLabor = values[9];
                        database.Add(cid.modelNum, cid);
                    }
                }
            }
            if (database.ContainsKey(methodModel))
            {
                CheckInData cid = database[methodModel];
                valveStem = cid.valveStem;
                oRing1 = cid.oRing1;
                oRing2 = cid.oRing2;
                size = cid.size;
                chemical = cid.chemical;
                htYear = cid.htYear;
                sixYear = cid.sixYear;
                rbCollar = cid.rbCollar;
                rbCT = cid.rbCT;
                extraParts = cid.extraParts;
                extraLabor = cid.extraLabor;
                validModel = true;
            }
            else
            {
                string path = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "/Check-In Templates/";
                DirectoryInfo di = Directory.CreateDirectory(path);
                string fileName = path + methodModel + " template.xml";
                if (File.Exists(fileName))
                {
                    TemplateBuilder loadedTemplate = Deserialize(fileName);

                    valveStem = loadedTemplate.templateVS;
                    oRing1 = loadedTemplate.templateOR1;
                    oRing2 = loadedTemplate.templateOR2;
                    size = loadedTemplate.templateSize;
                    chemical = loadedTemplate.templateChem;
                    htYear = loadedTemplate.templateHTYear;
                    sixYear = loadedTemplate.template6YRYear;
                    rbCollar = loadedTemplate.templateCollar;
                    rbCT = loadedTemplate.templateCT;
                    extraParts = loadedTemplate.templateExParts;
                    extraLabor = loadedTemplate.templateExLabor;
                    validModel = true;
                }
                else if ((File.Exists(fileName) == false) && (methodModel != "blank"))
                {

                    MessageBox.Show("Please add a new Extinguisher Template. Model number as searched: " + lastModelNumStr, "Model # Not Found", MessageBoxButtons.OK, MessageBoxIcon.Hand);
                    validModel = false;
                }
            }
            if (globalExLaborBool == true)
            {
                extraLabor = globalExtraLabor;
                globalExLaborBool = false;
            }
            if (globalExPartsBool == true)
            {
                extraParts = globalExtraParts;
                globalExPartsBool = false;
            }
        }

        // Designer.
        private System.ComponentModel.IContainer components = null;

        // Dispose of resources.
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        // Clear text boxes and return focus to the first box on enter keypress.
        private void Control_ReturnFocus(object sender, KeyEventArgs e)
        {
            if ((e.KeyCode == Keys.Enter) || (e.KeyCode == Keys.Return || e.KeyCode == Keys.Tab))
            {
                if (persistExtgInfo == true)
                {
                    yearManu.Clear();
                    walkinTakeback.Clear();
                    pressurized.Clear();
                    serialNum.Clear();

                    indivExParts.Clear();
                    indivExLabor.Clear();
                    e.Handled = true;
                    e.SuppressKeyPress = true;
                    if (condemnedCheck.Checked == true)
                    {
                        condemnedCheck.Checked = false;
                        globalExtraLabor = "";
                    }

                    serialNum.Focus();
                }
                else
                {
                    modelNum.Clear();
                    serialNum.Clear();
                    yearManu.Clear();
                    walkinTakeback.Clear();
                    pressurized.Clear();

                    indivExParts.Clear();
                    indivExLabor.Clear();
                    e.Handled = true;
                    e.SuppressKeyPress = true;
                    if (condemnedCheck.Checked == true)
                    {
                        condemnedCheck.Checked = false;
                        globalExtraLabor = "";
                    }
                    modelNum.Focus();
                }
            }
        }

        /* Same as above, but without keypress.
         * Future plans: Pretty sure there's no reason to have both of these be full methods. Either have Control_ReturnFocus call ReturnFocus, or completely delete one of them.
         */
        private void ReturnFocus()
        {
            ExcelReadoutRefresh();
            if (persistExtgInfo == true)
            {
                yearManu.Clear();
                walkinTakeback.Clear();
                pressurized.Clear();
                serialNum.Clear();

                indivExParts.Clear();
                indivExLabor.Clear();
                if (condemnedCheck.Checked == true)
                {
                    condemnedCheck.Checked = false;
                    globalExtraLabor = "";
                }

                serialNum.Focus();
            }
            else
            {
                modelNum.Clear();
                serialNum.Clear();
                yearManu.Clear();
                walkinTakeback.Clear();
                pressurized.Clear();

                indivExParts.Clear();
                indivExLabor.Clear();
                if (condemnedCheck.Checked == true)
                {
                    condemnedCheck.Checked = false;
                    globalExtraLabor = "";
                }
                modelNum.Focus();
            }
        }

        // Move focus to the next control.
        private void Control_KeyUp(object sender, KeyEventArgs e)
        {
            if ((e.KeyCode == Keys.Enter) || (e.KeyCode == Keys.Return || e.KeyCode == Keys.Tab))
            {
                this.SelectNextControl((System.Windows.Forms.Control)sender, true, true, true, true);
                e.Handled = true;
                e.SuppressKeyPress = true;
            }
        }

        /* Calls the serializer to deserialize the Template objects from the template xmls.
         * Future plans: This will be going out the window when I change Template files to .csv rows.
         */ 
        public static TemplateBuilder Deserialize(string fileName)
        {
            using (var stream = System.IO.File.OpenRead(fileName))
            {
                var serializer = new XmlSerializer(typeof(TemplateBuilder));
                return serializer.Deserialize(stream) as TemplateBuilder;
            }
        }

        // Manually submits the extinguisher record to the monthly sheet. Only runs when Submit button is clicked.
        public void SubmitButton()
        {
            int recordYearManu = Convert.ToInt32(recordYearManuStr);
            ExtgRecord record = new ExtgRecord(recordModelNum, recordSerialNum, recordYearManu, recordWITB, recordPressurized);

            if (recordModelNum != "blank")
            {
                RecordBuilder(recordModelNum, recordSerialNum, recordYearManu, recordWITB, recordPressurized);
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                string month = DateTime.Now.ToString("MMMM");
                int cYear = DateTime.Now.Year;
                if (validModel == true)
                {
                    //Write the new record to the Excel sheet.
                    string path = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                    string dirPath = path + "/Check-In Library/";
                    DirectoryInfo di = Directory.CreateDirectory(dirPath);
                    string fullPath = dirPath + month + " " + cYear + " Check-In.xlsx";
                    FileInfo excelFile = new FileInfo(fullPath);
                    using (ExcelPackage excel = new ExcelPackage(excelFile))
                    {
                        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                        if (SheetExist(fullPath, "Worksheet1") == false)
                        {
                            excel.Workbook.Worksheets.Add("Worksheet1");
                            List<string[]> headerRow = new List<string[]>()
                                {
                                    new string[] { "Model #", "Serial #", "Year of Manufacture", "Walk-In or Takeback", "Pressurized?", "Valve",
                                        "O-Ring 1", "O-Ring 2","Size", "Chemical", "Extra Parts", "Extra Labor", "Collar", "CT Label", "Customer","Town",
                                        "Date Arrived", "Claim Tag #", "Extg. # in Order", "Order Total","Current Year","Tally","Work Needed", "RC or Recondition"}
                                };
                            headerRange = "A1:" + Char.ConvertFromUtf32(headerRow[0].Length + 64) + "1";
                            var excelWorksheet = excel.Workbook.Worksheets["Worksheet1"];
                            excelWorksheet.Cells[headerRange].LoadFromArrays(headerRow);
                            excelWorksheet.Cells[headerRange].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            excelWorksheet.Cells[headerRange].Style.Font.Size = 14;
                            excelWorksheet.Cells[headerRange].Style.Font.Bold = true;
                            excelWorksheet.Cells[headerRange].AutoFitColumns();
                            var rowCnt = excelWorksheet.Dimension.End.Row;
                            int rowCount = rowCnt + 1;
                            var colCnt = excelWorksheet.Dimension.End.Column;
                            int colCount = colCnt;

                            if (rbCollar == true)
                            {
                                tagSealCollar = "1T, 1S, 1C";
                            }
                            else if (rbCollar == false)
                            {
                                tagSealCollar = "1T, 1S";
                            }

                            if (rbCT == true)
                            {
                                ctLabel = "CT Label";
                            }
                            else if (rbCT == false)
                            {
                                ctLabel = "";
                            }

                            excelWorksheet.Column(21).Hidden = true;

                            if ((currentYear - rbYear) >= htYear)
                            {
                                workNeeded = "HT";
                            }
                            else if (((currentYear - rbYear) >= sixYear) && ((currentYear - rbYear) < htYear))
                            {
                                workNeeded = "6YR";
                            }
                            else if (((currentYear - rbYear) < sixYear) && (rbPSI == true))
                            {
                                workNeeded = "INSP";
                            }
                            if ((rbPSI == true) && (workNeeded != "INSP"))
                            {
                                rechRecon = "Recondition";
                            }
                            else if (rbPSI == false)
                            {
                                rechRecon = "Recharge";
                            }
                            else
                            {
                                rechRecon = "INSP";
                            }

                            string rbPressureString;
                            if (rbPSI == true)
                            {
                                rbPressureString = "Yes";
                            } else
                            {
                                rbPressureString = "No";
                            }

                            var Records = new[]
                             {
                                    new {
                                model =  rbModel,
                                ser = record.clSerialNum,
                                year = rbYear,
                                wiTB = finalWITB,
                                PSI = rbPressureString,
                                vs = valveStem,
                                oR1 = oRing1,
                                oR2 = oRing2,
                                si = size,
                                chem = chemical,
                                exPart = extraParts,
                                exLabor = extraLabor,
                                collar = tagSealCollar,
                                CT = ctLabel,
                                HT = htYear,
                                sixYR = sixYear,
                                curYear = currentYear,
                                workNeeded,
                                rechRecon,
                                        }
                            };
                            foreach (var data in Records)
                            {
                                excelWorksheet.Cells[rowCount, 1].Value = data.model;
                                excelWorksheet.Cells[rowCount, 2].Value = data.ser;
                                excelWorksheet.Cells[rowCount, 3].Value = data.year;
                                excelWorksheet.Cells[rowCount, 4].Value = data.wiTB;
                                excelWorksheet.Cells[rowCount, 5].Value = data.PSI;
                                excelWorksheet.Cells[rowCount, 6].Value = data.vs;
                                excelWorksheet.Cells[rowCount, 7].Value = data.oR1;
                                excelWorksheet.Cells[rowCount, 8].Value = data.oR2;
                                excelWorksheet.Cells[rowCount, 9].Value = data.si;
                                excelWorksheet.Cells[rowCount, 10].Value = data.chem;
                                excelWorksheet.Cells[rowCount, 11].Value = data.exPart;
                                excelWorksheet.Cells[rowCount, 12].Value = data.exLabor;
                                excelWorksheet.Cells[rowCount, 13].Value = data.collar;
                                excelWorksheet.Cells[rowCount, 14].Value = data.CT;
                                excelWorksheet.Cells[rowCount, 21].Value = data.curYear;
                                excelWorksheet.Cells[rowCount, 23].Value = data.workNeeded;
                                excelWorksheet.Cells[rowCount, 24].Value = data.rechRecon;
                            }
                            excel.SaveAs(excelFile);
                            ReturnFocus();
                        }
                        else
                        {
                            ExcelWorksheet excelWorksheet = excel.Workbook.Worksheets.First();
                            var rowCnt = excelWorksheet.Dimension.End.Row;
                            int rowCount = rowCnt + 1;
                            var colCnt = excelWorksheet.Dimension.End.Column;
                            int colCount = colCnt;
                            excelWorksheet.Column(21).Hidden = true;
                            if (rbCollar == true)
                            {
                                tagSealCollar = "1T, 1S, 1C";
                            }
                            else if (rbCollar == false)
                            {
                                tagSealCollar = "1T, 1S";
                            }
                            if (rbCT == true)
                            {
                                ctLabel = "CT Label";
                            }
                            else if (rbCT == false)
                            {
                                ctLabel = "";
                            }
                            if ((currentYear - rbYear) >= htYear)
                            {
                                workNeeded = "HT";
                            }
                            else if (((currentYear - rbYear) >= sixYear) && ((currentYear - rbYear) < htYear))
                            {
                                workNeeded = "6YR";
                            }
                            else if (((currentYear - rbYear) < sixYear) && (rbPSI == true))
                            {
                                workNeeded = "INSP";
                            }
                            if ((rbPSI == true) && (workNeeded != "INSP"))
                            {
                                rechRecon = "Recondition";
                            }
                            else if (rbPSI == false)
                            {
                                rechRecon = "Recharge";
                            }
                            else
                            {
                                rechRecon = "INSP";
                            }
                            string rbPressureString;
                            if (rbPSI == true)
                            {
                                rbPressureString = "Yes";
                            }
                            else
                            {
                                rbPressureString = "No";
                            }

                            var Records = new[]
                            {
                                    new {
                                model =  rbModel,
                                ser = record.clSerialNum,
                                year = rbYear,
                                wiTB = finalWITB,
                                PSI = rbPressureString,
                                vs = valveStem,
                                oR1 = oRing1,
                                oR2 = oRing2,
                                si = size,
                                chem = chemical,
                                exPart = extraParts,
                                exLabor = extraLabor,
                                collar = tagSealCollar,
                                CT = ctLabel,
                                HT = htYear,
                                sixYR = sixYear,
                                curYear = currentYear,
                                workNeeded,
                                rechRecon,
                                    }
                                };
                            foreach (var data in Records)
                            {
                                excelWorksheet.Cells[rowCount, 1].Value = data.model;
                                excelWorksheet.Cells[rowCount, 2].Value = data.ser;
                                excelWorksheet.Cells[rowCount, 3].Value = data.year;
                                excelWorksheet.Cells[rowCount, 4].Value = data.wiTB;
                                excelWorksheet.Cells[rowCount, 5].Value = data.PSI;
                                excelWorksheet.Cells[rowCount, 6].Value = data.vs;
                                excelWorksheet.Cells[rowCount, 7].Value = data.oR1;
                                excelWorksheet.Cells[rowCount, 8].Value = data.oR2;
                                excelWorksheet.Cells[rowCount, 9].Value = data.si;
                                excelWorksheet.Cells[rowCount, 10].Value = data.chem;
                                excelWorksheet.Cells[rowCount, 11].Value = data.exPart;
                                excelWorksheet.Cells[rowCount, 12].Value = data.exLabor;
                                excelWorksheet.Cells[rowCount, 13].Value = data.collar;
                                excelWorksheet.Cells[rowCount, 14].Value = data.CT;
                                excelWorksheet.Cells[rowCount, 21].Value = data.curYear;
                                excelWorksheet.Cells[rowCount, 23].Value = data.workNeeded;
                                excelWorksheet.Cells[rowCount, 24].Value = data.rechRecon;
                            }
                            excel.SaveAs(excelFile);
                            ReturnFocus();
                        }
                    }
                }
            }
            else if (recordModelNum == "blank")
            {
                ReturnFocus();
            }
        }

        // Automatically saves the extinguisher record to the monthly sheet on Enter keypress when the user is at the final text box of the average record.
        private void Control_SaveRecord(object sender, KeyEventArgs e)
        {
            if ((e.KeyCode == Keys.Enter) || (e.KeyCode == Keys.Return) || (e.KeyCode == Keys.Tab))
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                //This section of code creates a new Extinguisher Record to add to the Excel file.
                int recordYearManu = Convert.ToInt32(recordYearManuStr);
                ExtgRecord record = new ExtgRecord(recordModelNum, recordSerialNum, recordYearManu, recordWITB, recordPressurized);

                if (recordModelNum != "blank")
                {
                    RecordBuilder(recordModelNum, recordSerialNum, recordYearManu, recordWITB, recordPressurized);

                    //Below is the section of code dedicated to creating and saving an Excel workbook for your records.

                    string month = DateTime.Now.ToString("MMMM");
                    int cYear = DateTime.Now.Year;
                    if (validModel == true)
                    {
                        //Write the new record to the Excel sheet.
                        string path = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                        string dirPath = path + "/Check-In Library/";
                        DirectoryInfo di = Directory.CreateDirectory(dirPath);
                        string fullPath = dirPath + month + " " + cYear + " Check-In.xlsx";
                        FileInfo excelFile = new FileInfo(fullPath);
                        using (ExcelPackage excel = new ExcelPackage(excelFile))
                        {
                            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                            if (SheetExist(fullPath, "Worksheet1") == false)
                            {
                                excel.Workbook.Worksheets.Add("Worksheet1");
                                List<string[]> headerRow = new List<string[]>()
                                {
                                    new string[] { "Model #", "Serial #", "Year of Manufacture", "Walk-In or Takeback", "Pressurized?", "Valve",
                                        "O-Ring 1", "O-Ring 2","Size", "Chemical", "Extra Parts", "Extra Labor", "Collar", "CT Label", "Customer","Town",
                                        "Date Arrived", "Claim Tag #", "Extg. # in Order", "Order Total","Current Year","Tally","Work Needed", "RC or Recondition"}
                                };

                                headerRange = "A1:" + Char.ConvertFromUtf32(headerRow[0].Length + 64) + "1";

                                var excelWorksheet = excel.Workbook.Worksheets["Worksheet1"];

                                excelWorksheet.Cells[headerRange].LoadFromArrays(headerRow);
                                excelWorksheet.Cells[headerRange].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                excelWorksheet.Cells[headerRange].Style.Font.Size = 14;
                                excelWorksheet.Cells[headerRange].Style.Font.Bold = true;
                                excelWorksheet.Cells[headerRange].AutoFitColumns();


                                var rowCnt = excelWorksheet.Dimension.End.Row;
                                int rowCount = rowCnt + 1;
                                var colCnt = excelWorksheet.Dimension.End.Column;
                                int colCount = colCnt;

                                if (rbCollar == true)
                                {
                                    tagSealCollar = "1T, 1S, 1C";
                                }
                                else if (rbCollar == false)
                                {
                                    tagSealCollar = "1T, 1S";
                                }

                                if (rbCT == true)
                                {
                                    ctLabel = "CT Label";
                                }
                                else if (rbCT == false)
                                {
                                    ctLabel = "";
                                }

                                excelWorksheet.Column(21).Hidden = true;

                                if ((currentYear - rbYear) >= htYear)
                                {
                                    workNeeded = "HT";
                                }
                                else if (((currentYear - rbYear) >= sixYear) && ((currentYear - rbYear) < htYear))
                                {
                                    workNeeded = "6YR";
                                }
                                else if (((currentYear - rbYear) < sixYear) && (rbPSI == true))
                                {
                                    workNeeded = "INSP";
                                }
                                if ((rbPSI == true) && (workNeeded != "INSP"))
                                {
                                    rechRecon = "Recondition";
                                }
                                else if (rbPSI == false)
                                {
                                    rechRecon = "Recharge";
                                }
                                else
                                {
                                    rechRecon = "INSP";
                                }

                                string rbPressureString;
                                if (rbPSI == true)
                                {
                                    rbPressureString = "Yes";
                                }
                                else
                                {
                                    rbPressureString = "No";
                                }

                                var Records = new[]
                                 {
                                    new {
                                model =  rbModel,
                                ser = record.clSerialNum,
                                year = rbYear,
                                wiTB = finalWITB,
                                PSI = rbPressureString,
                                vs = valveStem,
                                oR1 = oRing1,
                                oR2 = oRing2,
                                si = size,
                                chem = chemical,
                                exPart = extraParts,
                                exLabor = extraLabor,
                                collar = tagSealCollar,
                                CT = ctLabel,
                                HT = htYear,
                                sixYR = sixYear,
                                curYear = currentYear,
                                workNeeded,
                                rechRecon,


                                        }
                            };



                                foreach (var data in Records)
                                {
                                    excelWorksheet.Cells[rowCount, 1].Value = data.model;
                                    excelWorksheet.Cells[rowCount, 2].Value = data.ser;
                                    excelWorksheet.Cells[rowCount, 3].Value = data.year;
                                    excelWorksheet.Cells[rowCount, 4].Value = data.wiTB;
                                    excelWorksheet.Cells[rowCount, 5].Value = data.PSI;
                                    excelWorksheet.Cells[rowCount, 6].Value = data.vs;
                                    excelWorksheet.Cells[rowCount, 7].Value = data.oR1;
                                    excelWorksheet.Cells[rowCount, 8].Value = data.oR2;
                                    excelWorksheet.Cells[rowCount, 9].Value = data.si;
                                    excelWorksheet.Cells[rowCount, 10].Value = data.chem;
                                    excelWorksheet.Cells[rowCount, 11].Value = data.exPart;
                                    excelWorksheet.Cells[rowCount, 12].Value = data.exLabor;
                                    excelWorksheet.Cells[rowCount, 13].Value = data.collar;
                                    excelWorksheet.Cells[rowCount, 14].Value = data.CT;
                                    excelWorksheet.Cells[rowCount, 21].Value = data.curYear;
                                    excelWorksheet.Cells[rowCount, 23].Value = data.workNeeded;
                                    excelWorksheet.Cells[rowCount, 24].Value = data.rechRecon;
                                }
                                excel.SaveAs(excelFile);
                                ReturnFocus();
                            }
                            else
                            {
                                ExcelWorksheet excelWorksheet = excel.Workbook.Worksheets.First();
                                var rowCnt = excelWorksheet.Dimension.End.Row;
                                int rowCount = rowCnt + 1;
                                var colCnt = excelWorksheet.Dimension.End.Column;
                                int colCount = colCnt;
                                excelWorksheet.Column(21).Hidden = true;
                                if (rbCollar == true)
                                {
                                    tagSealCollar = "1T, 1S, 1C";
                                }
                                else if (rbCollar == false)
                                {
                                    tagSealCollar = "1T, 1S";
                                }

                                if (rbCT == true)
                                {
                                    ctLabel = "CT Label";
                                }
                                else if (rbCT == false)
                                {
                                    ctLabel = "";
                                }

                                if ((currentYear - rbYear) >= htYear)
                                {
                                    workNeeded = "HT";
                                }
                                else if (((currentYear - rbYear) >= sixYear) && ((currentYear - rbYear) < htYear))
                                {
                                    workNeeded = "6YR";
                                }
                                else if (((currentYear - rbYear) < sixYear) && (rbPSI == true))
                                {
                                    workNeeded = "INSP";
                                }
                                if ((rbPSI == true) && (workNeeded != "INSP"))
                                {
                                    rechRecon = "Recondition";
                                }
                                else if (rbPSI == false)
                                {
                                    rechRecon = "Recharge";
                                }
                                else
                                {
                                    rechRecon = "INSP";
                                }

                                string rbPressureString;
                                if (rbPSI == true)
                                {
                                    rbPressureString = "Yes";
                                }
                                else
                                {
                                    rbPressureString = "No";
                                }

                                var Records = new[]
                                {
                                    new {
                                model =  rbModel,
                                ser = record.clSerialNum,
                                year = rbYear,
                                wiTB = finalWITB,
                                PSI = rbPressureString,
                                vs = valveStem,
                                oR1 = oRing1,
                                oR2 = oRing2,
                                si = size,
                                chem = chemical,
                                exPart = extraParts,
                                exLabor = extraLabor,
                                collar = tagSealCollar,
                                CT = ctLabel,
                                HT = htYear,
                                sixYR = sixYear,
                                curYear = currentYear,
                                workNeeded,
                                rechRecon,
                                    }
                                };
                                foreach (var data in Records)
                                {
                                    excelWorksheet.Cells[rowCount, 1].Value = data.model;
                                    excelWorksheet.Cells[rowCount, 2].Value = data.ser;
                                    excelWorksheet.Cells[rowCount, 3].Value = data.year;
                                    excelWorksheet.Cells[rowCount, 4].Value = data.wiTB;
                                    excelWorksheet.Cells[rowCount, 5].Value = data.PSI;
                                    excelWorksheet.Cells[rowCount, 6].Value = data.vs;
                                    excelWorksheet.Cells[rowCount, 7].Value = data.oR1;
                                    excelWorksheet.Cells[rowCount, 8].Value = data.oR2;
                                    excelWorksheet.Cells[rowCount, 9].Value = data.si;
                                    excelWorksheet.Cells[rowCount, 10].Value = data.chem;
                                    excelWorksheet.Cells[rowCount, 11].Value = data.exPart;
                                    excelWorksheet.Cells[rowCount, 12].Value = data.exLabor;
                                    excelWorksheet.Cells[rowCount, 13].Value = data.collar;
                                    excelWorksheet.Cells[rowCount, 14].Value = data.CT;
                                    excelWorksheet.Cells[rowCount, 21].Value = data.curYear;
                                    excelWorksheet.Cells[rowCount, 23].Value = data.workNeeded;
                                    excelWorksheet.Cells[rowCount, 24].Value = data.rechRecon;
                                }
                                excel.SaveAs(excelFile);
                                ReturnFocus();
                            }
                        }
                    }
                }
                else if (recordModelNum == "blank")
                {
                    ReturnFocus();
                }
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
            this.components = new System.ComponentModel.Container();
            this.tableLayoutPanel1 = new System.Windows.Forms.TableLayoutPanel();
            this.flowLayoutPanel1 = new System.Windows.Forms.FlowLayoutPanel();
            this.Form2Button = new System.Windows.Forms.Button();
            this.custInfoButton = new System.Windows.Forms.Button();
            this.custDatabaseScreenButton = new System.Windows.Forms.Button();
            this.button3 = new System.Windows.Forms.Button();
            this.button1 = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.panel1 = new System.Windows.Forms.Panel();
            this.label6 = new System.Windows.Forms.Label();
            this.persistInfoBox = new System.Windows.Forms.CheckBox();
            this.condemnedCheck = new System.Windows.Forms.CheckBox();
            this.checkInSubmit = new System.Windows.Forms.Button();
            this.extraLaborCheck = new System.Windows.Forms.CheckBox();
            this.extraPartsCheck = new System.Windows.Forms.CheckBox();
            this.indivExLabor = new System.Windows.Forms.TextBox();
            this.indivExParts = new System.Windows.Forms.TextBox();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.pressurized = new System.Windows.Forms.TextBox();
            this.NoTemplateWarning = new System.Windows.Forms.Label();
            this.label11 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.walkinTakeback = new System.Windows.Forms.TextBox();
            this.yearManu = new System.Windows.Forms.TextBox();
            this.serialNum = new System.Windows.Forms.TextBox();
            this.modelNum = new System.Windows.Forms.TextBox();
            this.bindingSource1 = new System.Windows.Forms.BindingSource(this.components);
            this.tableLayoutPanel1.SuspendLayout();
            this.flowLayoutPanel1.SuspendLayout();
            this.panel1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.bindingSource1)).BeginInit();
            this.SuspendLayout();
            // 
            // tableLayoutPanel1
            // 
            this.tableLayoutPanel1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.tableLayoutPanel1.ColumnCount = 1;
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel1.Controls.Add(this.flowLayoutPanel1, 0, 1);
            this.tableLayoutPanel1.Controls.Add(this.panel1, 0, 0);
            this.tableLayoutPanel1.Location = new System.Drawing.Point(0, 0);
            this.tableLayoutPanel1.Name = "tableLayoutPanel1";
            this.tableLayoutPanel1.RowCount = 2;
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 89.47369F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 10.52632F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 20F));
            this.tableLayoutPanel1.Size = new System.Drawing.Size(815, 473);
            this.tableLayoutPanel1.TabIndex = 0;
            // 
            // flowLayoutPanel1
            // 
            this.flowLayoutPanel1.AutoSize = true;
            this.flowLayoutPanel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.flowLayoutPanel1.Controls.Add(this.Form2Button);
            this.flowLayoutPanel1.Controls.Add(this.custInfoButton);
            this.flowLayoutPanel1.Controls.Add(this.custDatabaseScreenButton);
            this.flowLayoutPanel1.Controls.Add(this.button3);
            this.flowLayoutPanel1.Controls.Add(this.button1);
            this.flowLayoutPanel1.Controls.Add(this.button2);
            this.flowLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.flowLayoutPanel1.Location = new System.Drawing.Point(3, 426);
            this.flowLayoutPanel1.Name = "flowLayoutPanel1";
            this.flowLayoutPanel1.Size = new System.Drawing.Size(809, 44);
            this.flowLayoutPanel1.TabIndex = 3;
            // 
            // Form2Button
            // 
            this.Form2Button.Anchor = System.Windows.Forms.AnchorStyles.Bottom;
            this.Form2Button.AutoSize = true;
            this.Form2Button.Location = new System.Drawing.Point(3, 3);
            this.Form2Button.Name = "Form2Button";
            this.Form2Button.Size = new System.Drawing.Size(103, 23);
            this.Form2Button.TabIndex = 5;
            this.Form2Button.Text = "&Add New Model #";
            this.Form2Button.UseVisualStyleBackColor = true;
            this.Form2Button.Click += new System.EventHandler(this.button4_Click);
            // 
            // custInfoButton
            // 
            this.custInfoButton.Anchor = System.Windows.Forms.AnchorStyles.Bottom;
            this.custInfoButton.AutoSize = true;
            this.custInfoButton.Location = new System.Drawing.Point(112, 3);
            this.custInfoButton.Name = "custInfoButton";
            this.custInfoButton.Size = new System.Drawing.Size(110, 23);
            this.custInfoButton.TabIndex = 6;
            this.custInfoButton.Text = "&Enter Customer Info";
            this.custInfoButton.UseVisualStyleBackColor = true;
            this.custInfoButton.Click += new System.EventHandler(this.custInfoButton_Click);
            // 
            // custDatabaseScreenButton
            // 
            this.custDatabaseScreenButton.Location = new System.Drawing.Point(228, 3);
            this.custDatabaseScreenButton.Name = "custDatabaseScreenButton";
            this.custDatabaseScreenButton.Size = new System.Drawing.Size(120, 23);
            this.custDatabaseScreenButton.TabIndex = 10;
            this.custDatabaseScreenButton.Text = "&Build Cust. Database";
            this.custDatabaseScreenButton.UseVisualStyleBackColor = true;
            this.custDatabaseScreenButton.Click += new System.EventHandler(this.custDatabaseScreenButton_Click);
            // 
            // button3
            // 
            this.button3.Anchor = System.Windows.Forms.AnchorStyles.Bottom;
            this.button3.AutoSize = true;
            this.button3.Location = new System.Drawing.Point(354, 3);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(166, 23);
            this.button3.TabIndex = 9;
            this.button3.Text = "&Calculate Formulas in Database";
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Click += new System.EventHandler(this.button3_Click_1);
            // 
            // button1
            // 
            this.button1.Anchor = System.Windows.Forms.AnchorStyles.Bottom;
            this.button1.AutoSize = true;
            this.button1.Location = new System.Drawing.Point(526, 3);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(127, 23);
            this.button1.TabIndex = 7;
            this.button1.Text = "&Refresh Excel Readout";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // button2
            // 
            this.button2.Anchor = System.Windows.Forms.AnchorStyles.Bottom;
            this.button2.AutoSize = true;
            this.button2.Location = new System.Drawing.Point(659, 3);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(132, 23);
            this.button2.TabIndex = 8;
            this.button2.Text = "&Open Monthly Database";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // panel1
            // 
            this.panel1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.panel1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.panel1.Controls.Add(this.label6);
            this.panel1.Controls.Add(this.persistInfoBox);
            this.panel1.Controls.Add(this.condemnedCheck);
            this.panel1.Controls.Add(this.checkInSubmit);
            this.panel1.Controls.Add(this.extraLaborCheck);
            this.panel1.Controls.Add(this.extraPartsCheck);
            this.panel1.Controls.Add(this.indivExLabor);
            this.panel1.Controls.Add(this.indivExParts);
            this.panel1.Controls.Add(this.dataGridView1);
            this.panel1.Controls.Add(this.pressurized);
            this.panel1.Controls.Add(this.NoTemplateWarning);
            this.panel1.Controls.Add(this.label11);
            this.panel1.Controls.Add(this.label5);
            this.panel1.Controls.Add(this.label4);
            this.panel1.Controls.Add(this.label3);
            this.panel1.Controls.Add(this.label2);
            this.panel1.Controls.Add(this.label1);
            this.panel1.Controls.Add(this.walkinTakeback);
            this.panel1.Controls.Add(this.yearManu);
            this.panel1.Controls.Add(this.serialNum);
            this.panel1.Controls.Add(this.modelNum);
            this.panel1.Location = new System.Drawing.Point(3, 3);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(809, 417);
            this.panel1.TabIndex = 4;
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(-1, 125);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(356, 13);
            this.label6.TabIndex = 27;
            this.label6.Text = "Reminder: Don\'t forget the - in the serial number when entering it manually!";
            // 
            // persistInfoBox
            // 
            this.persistInfoBox.Anchor = System.Windows.Forms.AnchorStyles.Top;
            this.persistInfoBox.AutoSize = true;
            this.persistInfoBox.Location = new System.Drawing.Point(347, 30);
            this.persistInfoBox.Name = "persistInfoBox";
            this.persistInfoBox.Size = new System.Drawing.Size(111, 17);
            this.persistInfoBox.TabIndex = 26;
            this.persistInfoBox.Text = "Persist Extg. Info?";
            this.persistInfoBox.UseVisualStyleBackColor = true;
            this.persistInfoBox.CheckedChanged += new System.EventHandler(this.persistInfoBox_CheckedChanged);
            // 
            // condemnedCheck
            // 
            this.condemnedCheck.Anchor = System.Windows.Forms.AnchorStyles.Top;
            this.condemnedCheck.AutoSize = true;
            this.condemnedCheck.Location = new System.Drawing.Point(364, 93);
            this.condemnedCheck.Name = "condemnedCheck";
            this.condemnedCheck.Size = new System.Drawing.Size(89, 17);
            this.condemnedCheck.TabIndex = 25;
            this.condemnedCheck.Text = "Condemned?";
            this.condemnedCheck.UseVisualStyleBackColor = true;
            this.condemnedCheck.CheckedChanged += new System.EventHandler(this.condemnedCheck_CheckedChanged);
            // 
            // checkInSubmit
            // 
            this.checkInSubmit.Anchor = System.Windows.Forms.AnchorStyles.Top;
            this.checkInSubmit.Location = new System.Drawing.Point(364, 120);
            this.checkInSubmit.Name = "checkInSubmit";
            this.checkInSubmit.Size = new System.Drawing.Size(75, 23);
            this.checkInSubmit.TabIndex = 24;
            this.checkInSubmit.Text = "&Submit";
            this.checkInSubmit.UseVisualStyleBackColor = true;
            this.checkInSubmit.Click += new System.EventHandler(this.checkInSubmit_Click);
            // 
            // extraLaborCheck
            // 
            this.extraLaborCheck.Anchor = System.Windows.Forms.AnchorStyles.Top;
            this.extraLaborCheck.AutoSize = true;
            this.extraLaborCheck.Location = new System.Drawing.Point(483, 104);
            this.extraLaborCheck.Name = "extraLaborCheck";
            this.extraLaborCheck.Size = new System.Drawing.Size(86, 17);
            this.extraLaborCheck.TabIndex = 23;
            this.extraLaborCheck.Text = "Extra Labor?";
            this.extraLaborCheck.UseVisualStyleBackColor = true;
            this.extraLaborCheck.CheckedChanged += new System.EventHandler(this.extraLaborCheck_CheckedChanged);
            // 
            // extraPartsCheck
            // 
            this.extraPartsCheck.Anchor = System.Windows.Forms.AnchorStyles.Top;
            this.extraPartsCheck.AutoSize = true;
            this.extraPartsCheck.Location = new System.Drawing.Point(93, 104);
            this.extraPartsCheck.Name = "extraPartsCheck";
            this.extraPartsCheck.Size = new System.Drawing.Size(83, 17);
            this.extraPartsCheck.TabIndex = 22;
            this.extraPartsCheck.Text = "Extra Parts?";
            this.extraPartsCheck.UseVisualStyleBackColor = true;
            this.extraPartsCheck.CheckedChanged += new System.EventHandler(this.extraPartsCheck_CheckedChanged);
            // 
            // indivExLabor
            // 
            this.indivExLabor.Anchor = System.Windows.Forms.AnchorStyles.Top;
            this.indivExLabor.Location = new System.Drawing.Point(613, 101);
            this.indivExLabor.Name = "indivExLabor";
            this.indivExLabor.Size = new System.Drawing.Size(100, 20);
            this.indivExLabor.TabIndex = 21;
            this.indivExLabor.TextChanged += new System.EventHandler(this.indivExLabor_TextChanged);
            // 
            // indivExParts
            // 
            this.indivExParts.Anchor = System.Windows.Forms.AnchorStyles.Top;
            this.indivExParts.Location = new System.Drawing.Point(223, 101);
            this.indivExParts.Name = "indivExParts";
            this.indivExParts.Size = new System.Drawing.Size(100, 20);
            this.indivExParts.TabIndex = 19;
            this.indivExParts.TextChanged += new System.EventHandler(this.indivExParts_TextChanged);
            // 
            // dataGridView1
            // 
            this.dataGridView1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Location = new System.Drawing.Point(0, 149);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.Size = new System.Drawing.Size(805, 264);
            this.dataGridView1.TabIndex = 17;
            // 
            // pressurized
            // 
            this.pressurized.Anchor = System.Windows.Forms.AnchorStyles.Top;
            this.pressurized.Location = new System.Drawing.Point(613, 66);
            this.pressurized.Name = "pressurized";
            this.pressurized.Size = new System.Drawing.Size(101, 20);
            this.pressurized.TabIndex = 4;
            this.pressurized.TextChanged += new System.EventHandler(this.pressurized_TextChanged);
            this.pressurized.Enter += new System.EventHandler(this.pressurized_Enter);
            this.pressurized.KeyDown += new System.Windows.Forms.KeyEventHandler(this.Control_SaveRecord);
            // 
            // NoTemplateWarning
            // 
            this.NoTemplateWarning.AutoSize = true;
            this.NoTemplateWarning.Location = new System.Drawing.Point(55, 260);
            this.NoTemplateWarning.Name = "NoTemplateWarning";
            this.NoTemplateWarning.Size = new System.Drawing.Size(0, 13);
            this.NoTemplateWarning.TabIndex = 16;
            // 
            // label11
            // 
            this.label11.Anchor = System.Windows.Forms.AnchorStyles.Top;
            this.label11.AutoSize = true;
            this.label11.Font = new System.Drawing.Font("Microsoft Sans Serif", 18F, System.Drawing.FontStyle.Bold);
            this.label11.Location = new System.Drawing.Point(88, 0);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(620, 29);
            this.label11.TabIndex = 15;
            this.label11.Text = "Please enter the extinguisher\'s check-in information.";
            // 
            // label5
            // 
            this.label5.Anchor = System.Windows.Forms.AnchorStyles.Top;
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(630, 50);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(67, 13);
            this.label5.TabIndex = 9;
            this.label5.Text = "Pressurized?";
            // 
            // label4
            // 
            this.label4.Anchor = System.Windows.Forms.AnchorStyles.Top;
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(480, 50);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(108, 13);
            this.label4.TabIndex = 8;
            this.label4.Text = "Walk-In or Takeback";
            // 
            // label3
            // 
            this.label3.Anchor = System.Windows.Forms.AnchorStyles.Top;
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(350, 50);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(108, 13);
            this.label3.TabIndex = 7;
            this.label3.Text = "Year of Last 6YR/HT";
            // 
            // label2
            // 
            this.label2.Anchor = System.Windows.Forms.AnchorStyles.Top;
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(236, 50);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(73, 13);
            this.label2.TabIndex = 6;
            this.label2.Text = "Serial Number";
            // 
            // label1
            // 
            this.label1.Anchor = System.Windows.Forms.AnchorStyles.Top;
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(104, 50);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(76, 13);
            this.label1.TabIndex = 5;
            this.label1.Text = "Model Number";
            // 
            // walkinTakeback
            // 
            this.walkinTakeback.Anchor = System.Windows.Forms.AnchorStyles.Top;
            this.walkinTakeback.Location = new System.Drawing.Point(483, 66);
            this.walkinTakeback.Name = "walkinTakeback";
            this.walkinTakeback.Size = new System.Drawing.Size(100, 20);
            this.walkinTakeback.TabIndex = 3;
            this.walkinTakeback.TextChanged += new System.EventHandler(this.walkinTakeback_TextChanged);
            this.walkinTakeback.Enter += new System.EventHandler(this.walkinTakeback_Enter);
            this.walkinTakeback.KeyDown += new System.Windows.Forms.KeyEventHandler(this.Control_KeyUp);
            // 
            // yearManu
            // 
            this.yearManu.Anchor = System.Windows.Forms.AnchorStyles.Top;
            this.yearManu.Location = new System.Drawing.Point(353, 66);
            this.yearManu.Name = "yearManu";
            this.yearManu.Size = new System.Drawing.Size(100, 20);
            this.yearManu.TabIndex = 2;
            this.yearManu.TextChanged += new System.EventHandler(this.yearManu_TextChanged);
            this.yearManu.Enter += new System.EventHandler(this.yearManu_Enter);
            this.yearManu.KeyDown += new System.Windows.Forms.KeyEventHandler(this.Control_KeyUp);
            // 
            // serialNum
            // 
            this.serialNum.Anchor = System.Windows.Forms.AnchorStyles.Top;
            this.serialNum.Location = new System.Drawing.Point(223, 66);
            this.serialNum.Name = "serialNum";
            this.serialNum.Size = new System.Drawing.Size(100, 20);
            this.serialNum.TabIndex = 1;
            this.serialNum.TextChanged += new System.EventHandler(this.serialNum_TextChanged);
            this.serialNum.Enter += new System.EventHandler(this.serialNum_Enter);
            this.serialNum.KeyDown += new System.Windows.Forms.KeyEventHandler(this.Control_KeyUp);
            // 
            // modelNum
            // 
            this.modelNum.Anchor = System.Windows.Forms.AnchorStyles.Top;
            this.modelNum.Location = new System.Drawing.Point(93, 66);
            this.modelNum.Name = "modelNum";
            this.modelNum.Size = new System.Drawing.Size(100, 20);
            this.modelNum.TabIndex = 0;
            this.modelNum.TextChanged += new System.EventHandler(this.ModelNum_TextChanged);
            this.modelNum.Enter += new System.EventHandler(this.modelNum_Enter);
            this.modelNum.KeyDown += new System.Windows.Forms.KeyEventHandler(this.Control_KeyUp);
            // 
            // CheckIn
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(815, 473);
            this.Controls.Add(this.tableLayoutPanel1);
            this.Name = "CheckIn";
            this.Text = "Check-In";
            this.WindowState = System.Windows.Forms.FormWindowState.Maximized;
            this.Shown += new System.EventHandler(this.CheckIn_Shown);
            this.tableLayoutPanel1.ResumeLayout(false);
            this.tableLayoutPanel1.PerformLayout();
            this.flowLayoutPanel1.ResumeLayout(false);
            this.flowLayoutPanel1.PerformLayout();
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.bindingSource1)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel1;
        private System.Windows.Forms.FlowLayoutPanel flowLayoutPanel1;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox walkinTakeback;
        private System.Windows.Forms.TextBox yearManu;
        private System.Windows.Forms.TextBox serialNum;
        private System.Windows.Forms.TextBox modelNum;
        private Button Form2Button;
        private Label label11;
        private Label NoTemplateWarning;
        private TextBox pressurized;
        private Button custInfoButton;
        private DataGridView dataGridView1;
        private BindingSource bindingSource1;
        private Button button1;
        private Button button2;
        private Button button3;
        private TextBox indivExLabor;
        private TextBox indivExParts;
        private CheckBox extraLaborCheck;
        private CheckBox extraPartsCheck;
        private Button checkInSubmit;
        private CheckBox condemnedCheck;
        private CheckBox persistInfoBox;
        private Button custDatabaseScreenButton;
        private Label label6;
    }
}

