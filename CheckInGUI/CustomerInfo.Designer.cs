using System;
using System.Windows.Forms;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Linq;
using System.IO;
using OfficeOpenXml;
using System.Collections.Generic;
using OfficeOpenXml.Style;
using System.Xml;
using System.Xml.Serialization;
using System.Data;
using System.Media;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Office2013.Excel;

namespace CheckInGUI
{
    partial class CustomerInfo
    {
        public static int emptyRowAddress;
        public static int emptyRecRowAddress;
        
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

        // Move focus to the next control on enter keypress.
        private void Control_KeyUp(object sender, KeyEventArgs e)
        {
            if ((e.KeyCode == Keys.Enter) || (e.KeyCode == Keys.Return || e.KeyCode == Keys.Tab))
            {
                this.SelectNextControl((System.Windows.Forms.Control)sender, true, true, true, true);
                e.Handled = true;
                e.SuppressKeyPress = true;
            }
        }

        // Refreshes the live-view of the Excel spreadsheet for the month.
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
                    System.Data.DataTable dt = ExcelPackageToDataTable(excel2);
                    dataGridView1.DataSource = dt;
                    dataGridView1.ClearSelection();

                    int nRowIndex = dataGridView1.Rows.Count - 1;

                    dataGridView1.Rows[nRowIndex].Selected = true;
                    dataGridView1.FirstDisplayedScrollingRowIndex = nRowIndex;

                }
            }
        }

        // Checks for the existence of a worksheet in a workbook.
        internal static bool SheetExist(string fullFilePath, string sheetName)
        {
            using (var package = new ExcelPackage(new FileInfo(fullFilePath)))
            {
                return package.Workbook.Worksheets.Any(sheet => sheet.Name == sheetName);
            }
        }

        // Converts the ExcelPackage with all data in it to a datatable for live-view.
        public static System.Data.DataTable ExcelPackageToDataTable(ExcelPackage excelPackage)
        {
            System.Data.DataTable dt = new System.Data.DataTable();
            ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets[0];

            //check if the worksheet is completely empty
            if (worksheet.Dimension == null)
            {
                return dt;
            }

            //create a list to hold the column names
            List<string> columnNames = new List<string>();

            //needed to keep track of empty column headers
            int currentColumn = 1;

            //loop all columns in the sheet and add them to the datatable
            foreach (var cell in worksheet.Cells[1, 1, 1, worksheet.Dimension.End.Column])
            {
                string columnName = cell.Text.Trim();

                //check if the previous header was empty and add it if it was
                if (cell.Start.Column != currentColumn)
                {
                    columnNames.Add("Header_" + currentColumn);
                    dt.Columns.Add("Header_" + currentColumn);
                    currentColumn++;
                }

                //add the column name to the list to count the duplicates
                columnNames.Add(columnName);

                //count the duplicate column names and make them unique to avoid the exception
                //A column named 'Name' already belongs to this DataTable
                int occurrences = columnNames.Count(x => x.Equals(columnName));
                if (occurrences > 1)
                {
                    columnName = columnName + "_" + occurrences;
                }

                //add the column to the datatable
                dt.Columns.Add(columnName);

                currentColumn++;
            }

            //start adding the contents of the excel file to the datatable
            for (int i = 2; i <= worksheet.Dimension.End.Row; i++)
            {
                var row = worksheet.Cells[i, 1, i, worksheet.Dimension.End.Column];
                DataRow newRow = dt.NewRow();

                //loop all cells in the row
                foreach (var cell in row)
                {
                    newRow[cell.Start.Column - 1] = cell.Text;
                }

                dt.Rows.Add(newRow);
            }
            return dt;
        }

        // Clears all text boxes and returns focus to the first text box.
        private void Control_ReturnFocus()
        {
            

            custName.Clear();
            dateRec.Clear();
            claimNum.Clear();
            custTown.Clear();

            custName.Focus();
            
        }

        // A special version of Control_ReturnFocus that only clears the ClaimNum box, in case two extinguishers belong to the same customer.
        private void SameCustomerFocus()
        {
            claimNum.Clear();
            claimNum.Focus();
        }
        
        // Manually adds the data to the existing spreadsheet and saves the spreadsheet on click of Submit button.
        private void Submit_SaveRecord()
        {
            //Below is the section of code dedicated to creating and saving an Excel workbook for your records.

            //Write the new record to the Excel sheet.
            string month = DateTime.Now.ToString("MMMM");
            int cYear = DateTime.Now.Year;

            //Write the new record to the Excel sheet.
            string path = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            string dirPath = path + "/Check-In Library/";
            DirectoryInfo di = Directory.CreateDirectory(dirPath);
            string fullPath = dirPath + month + " " + cYear + " Check-In.xlsx";
            FileInfo excelFile = new FileInfo(fullPath);
            {
                using (ExcelPackage excel = new ExcelPackage(excelFile))
                {

                    ExcelWorksheet worksheet = excel.Workbook.Worksheets.FirstOrDefault(x => x.Name == "Worksheet1");
                    List<string[]> headerRow = new List<string[]>()
                                {
                                    new string[] { "Model #", "Serial #", "Year of Manufacture", "Walk-In or Takeback", "Pressurized?", "Valve",
                                        "O-Ring 1", "O-Ring 2","Size", "Chemical", "Extra Parts", "Extra Labor", "Collar", "CT Label", "Customer","Town",
                                        "Date Arrived", "Claim Tag #", "Extg. # in Order", "Order Total","Current Year","Tally","Work Needed", "RC or Recondition"}
                                };

                    string headerRange = "A1:" + Char.ConvertFromUtf32(headerRow[0].Length + 64) + "1";
                    int colCount = 15;
                    bool emptyRow = false;
                    int rowCounter = 1;
                    bool emptyCell = false;
                    int recCounter = 1;


                    while (emptyRow == false)
                    {

                        if (worksheet.Cells[rowCounter, colCount].Value == null)
                        {

                            emptyRowAddress = rowCounter;

                            emptyRow = true;
                        }
                        else
                        {
                            rowCounter++;
                        }

                    }
                    while (emptyCell == false)
                    {

                        if (worksheet.Cells[recCounter, 1].Value == null)
                        {

                            emptyRecRowAddress = recCounter;

                            emptyCell = true;
                        }
                        else
                        {
                            recCounter++;
                        }
                    }

                    int loopCount = emptyRowAddress;
                    int loopRowCount = 1;
                    while (loopCount <= emptyRecRowAddress)
                    {
                        if (worksheet.Cells[loopRowCount, 1].Value != null)
                        {

                            var CustRecords = new[]
                            {
                                    new {
                                            cust = custNameData,
                                            date = dateRecData,
                                            claim = claimNumData,
                                            town = custTownData,
                                        }
                            };
                            if (emptyRowAddress == 2)
                            {
                                foreach (var data in CustRecords)
                                {
                                    worksheet.Cells[emptyRowAddress, 15].Value = data.cust;
                                    worksheet.Cells[emptyRowAddress, 16].Value = data.town;
                                    worksheet.Cells[emptyRowAddress, 17].Value = data.date;
                                    worksheet.Cells[emptyRowAddress, 18].Value = data.claim;
                                    worksheet.Cells[headerRange].AutoFitColumns();
                                }
                                foreach (var data in CustRecords)
                                {
                                    int cnt = 2;
                                    bool empty = false;
                                    int empRow = 0;
                                    while (empty == false)
                                    {

                                        if (worksheet.Cells[cnt, 1].Value == null)
                                        {

                                            empRow = cnt;

                                            empty = true;
                                        }
                                        else
                                        {
                                            cnt++;
                                        }
                                    }
                                    int decRowNum = empRow;
                                    int total = 1;
                                    int tally = 1;
                                    while (decRowNum > 1)
                                    {
                                        string decCust = worksheet.Cells[decRowNum, 15].ToString();
                                        string decTown = worksheet.Cells[decRowNum, 16].ToString();
                                        string decDate = worksheet.Cells[decRowNum, 17].ToString();
                                        bool sameCust = false;
                                        bool sameTown = false;
                                        bool sameDate = false;
                                        if (decCust == data.cust)
                                        {
                                            sameCust = true;
                                        }
                                        if (decTown == data.town)
                                        {
                                            sameTown = true;
                                        }
                                        if (decDate == data.date)
                                        {
                                            sameDate = true;
                                        }
                                        if ((sameCust == true) && (sameTown == true) && (sameDate == true))
                                        {
                                            total++;
                                        }
                                        worksheet.Cells[decRowNum, 20].Value = total;
                                        worksheet.Cells[headerRange].AutoFitColumns();
                                        if (total > 1)
                                        {
                                            int rowCount2 = 2;
                                            while (rowCount2 < empRow)
                                            {
                                                // This needs to do multiple things.
                                                // 1. Iterate through entire excel sheet.
                                                // 2. Compare a customer name to every single cell in column 15.
                                                // 3. Compare a customer town to every single cell in column 16.
                                                // 4. Compare a customer date to every single cell in column 17.
                                                // 5. If all three of those find a match on the same cell, then put the tally on that cell.
                                                // 6. Increment the counter past the previous match, look again, add the tally + 1 on that cell.
                                                // 7. Repeat #6 until no matches are found.
                                                bool custMatch = false;
                                                bool townMatch = false;
                                                bool dateMatch = false;
                                                string custText = worksheet.Cells[rowCount2, 15].Value.ToString();
                                                string townText = worksheet.Cells[rowCount2, 16].Value.ToString();
                                                string dateText = worksheet.Cells[rowCount2, 17].Value.ToString();
                                                if (custText == decCust)
                                                {
                                                    custMatch = true;
                                                }
                                                if (townText == decTown)
                                                {
                                                    townMatch = true;
                                                }
                                                if (dateText == decDate)
                                                {
                                                    dateMatch = true;
                                                }
                                                if ((custMatch == true) && (townMatch == true) && (dateMatch == true))
                                                {
                                                    worksheet.Cells[rowCount2, 19].Value = tally;
                                                    tally++;
                                                    bool tallyMatch = true;
                                                    while (tallyMatch == true)
                                                    {
                                                        string tallyCustText = worksheet.Cells[rowCount2, 15].Value.ToString();
                                                        string tallyTownText = worksheet.Cells[rowCount2, 16].Value.ToString();
                                                        string tallyDateText = worksheet.Cells[rowCount2, 17].Value.ToString();
                                                        bool custTally = false;
                                                        bool townTally = false;
                                                        bool dateTally = false;
                                                        if (tallyCustText == custText)
                                                        {
                                                            custTally = true;
                                                        }
                                                        if (tallyTownText == townText)
                                                        {
                                                            townTally = true;
                                                        }
                                                        if (tallyDateText == dateText)
                                                        {
                                                            dateTally = true;
                                                        }
                                                        if ((custTally == true) && (townTally == true) && (dateTally == true))
                                                        {
                                                            worksheet.Cells[rowCount2, 19].Value = tally;
                                                            worksheet.Cells[headerRange].AutoFitColumns();
                                                            if (tally == total)
                                                            {
                                                                tallyMatch = false;
                                                            }
                                                            tally++;
                                                            rowCount2++;
                                                        }
                                                    }
                                                }

                                            }
                                        }
                                        string decRowTally = worksheet.Cells[decRowNum, 19].Value.ToString();
                                        string decRowTotal = worksheet.Cells[decRowNum, 20].Value.ToString();
                                        string cell21 = decRowTally + " of " + decRowTotal;
                                        worksheet.Cells[decRowNum, 21].Value = cell21;
                                        worksheet.Cells[headerRange].AutoFitColumns();
                                    }
                                }
                            }
                            else
                            {
                                foreach (var data in CustRecords)
                                {
                                    worksheet.Cells[emptyRowAddress, 15].Value = data.cust;
                                    worksheet.Cells[emptyRowAddress, 16].Value = data.town;
                                    worksheet.Cells[emptyRowAddress, 17].Value = data.date;
                                    worksheet.Cells[emptyRowAddress, 18].Value = data.claim;
                                    worksheet.Cells[headerRange].AutoFitColumns();
                                    
                                }
                                foreach (var data in CustRecords)
                                {
                                    int cnt = 2;
                                    bool empty = false;
                                    int empRow = 0;
                                    while (empty == false)
                                    {

                                        if (worksheet.Cells[cnt, 1].Value == null)
                                        {

                                            empRow = cnt;

                                            empty = true;
                                        }
                                        else
                                        {
                                            cnt++;
                                        }
                                    }
                                    int decRowNum = empRow;
                                    int total = 1;
                                    int tally = 1;
                                    while (decRowNum > 1)
                                    {
                                        string decCust = worksheet.Cells[decRowNum, 15].ToString();
                                        string decTown = worksheet.Cells[decRowNum, 16].ToString();
                                        string decDate = worksheet.Cells[decRowNum, 17].ToString();
                                        bool sameCust = false;
                                        bool sameTown = false;
                                        bool sameDate = false;
                                        if (decCust == data.cust)
                                        {
                                            sameCust = true;
                                        }
                                        if (decTown == data.town)
                                        {
                                            sameTown = true;
                                        }
                                        if (decDate == data.date)
                                        {
                                            sameDate = true;
                                        }
                                        if ((sameCust == true) && (sameTown == true) && (sameDate == true))
                                        {
                                            total++;
                                        }
                                        worksheet.Cells[decRowNum, 20].Value = total;
                                        worksheet.Cells[headerRange].AutoFitColumns();
                                        if (total > 1)
                                        {
                                            int rowCount2 = 2;
                                            while (rowCount2 < empRow)
                                            {
                                                // This needs to do multiple things.
                                                // 1. Iterate through entire excel sheet.
                                                // 2. Compare a customer name to every single cell in column 15.
                                                // 3. Compare a customer town to every single cell in column 16.
                                                // 4. Compare a customer date to every single cell in column 17.
                                                // 5. If all three of those find a match on the same cell, then put the tally on that cell.
                                                // 6. Increment the counter past the previous match, look again, add the tally + 1 on that cell.
                                                // 7. Repeat #6 until no matches are found.
                                                bool custMatch = false;
                                                bool townMatch = false;
                                                bool dateMatch = false;
                                                string custText = worksheet.Cells[rowCount2, 15].Value.ToString();
                                                string townText = worksheet.Cells[rowCount2, 16].Value.ToString();
                                                string dateText = worksheet.Cells[rowCount2, 17].Value.ToString();
                                                if (custText == decCust)
                                                {
                                                    custMatch = true;
                                                }
                                                if (townText == decTown)
                                                {
                                                    townMatch = true;
                                                }
                                                if (dateText == decDate)
                                                {
                                                    dateMatch = true;
                                                }
                                                if ((custMatch == true) && (townMatch == true) && (dateMatch == true))
                                                {
                                                    worksheet.Cells[rowCount2, 19].Value = tally;
                                                    tally++;
                                                    bool tallyMatch = true;
                                                    while (tallyMatch == true)
                                                    {
                                                        string tallyCustText = worksheet.Cells[rowCount2, 15].Value.ToString();
                                                        string tallyTownText = worksheet.Cells[rowCount2, 16].Value.ToString();
                                                        string tallyDateText = worksheet.Cells[rowCount2, 17].Value.ToString();
                                                        bool custTally = false;
                                                        bool townTally = false;
                                                        bool dateTally = false;
                                                        if (tallyCustText == custText)
                                                        {
                                                            custTally = true;
                                                        }
                                                        if (tallyTownText == townText)
                                                        {
                                                            townTally = true;
                                                        }
                                                        if (tallyDateText == dateText)
                                                        {
                                                            dateTally = true;
                                                        }
                                                        if ((custTally == true) && (townTally == true) && (dateTally == true))
                                                        {
                                                            worksheet.Cells[rowCount2, 19].Value = tally;
                                                            worksheet.Cells[headerRange].AutoFitColumns();
                                                            if (tally == total)
                                                            {
                                                                tallyMatch = false;
                                                            }
                                                            tally++;
                                                            rowCount2++;
                                                        }
                                                    }
                                                }

                                            }
                                        }
                                        string decRowTally = worksheet.Cells[decRowNum, 19].Value.ToString();
                                        string decRowTotal = worksheet.Cells[decRowNum, 20].Value.ToString();
                                        string cell21 = decRowTally + " of " + decRowTotal;
                                        worksheet.Cells[decRowNum, 21].Value = cell21;
                                        worksheet.Cells[headerRange].AutoFitColumns();
                                    }
                                }                                                               
                            }
                        excel.SaveAs(excelFile);
                                loopCount++;
                                loopRowCount++;
                        }
                    }
                    if (worksheet.Cells[emptyRowAddress, 1].Value == null)
                    {
                        SystemSounds.Hand.Play();
                        MessageBox.Show("Customer Info has surpassed Extg Info", "OVERLOAD", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }

            }
            if (sameCust == false)
            {
                Control_ReturnFocus();
            }
            else
            {
                SameCustomerFocus();
            }
        ExcelReadoutRefresh();
        }

        // Automatically saves the data to the existing spreadsheet on enter keypress when the user reaches the final text box.
        private void Control_SaveRecord(object sender, KeyEventArgs e)
        {
            if ((e.KeyCode == Keys.Enter) || (e.KeyCode == Keys.Return || e.KeyCode == Keys.Tab))
            {
                e.Handled = e.SuppressKeyPress = true;
                //Below is the section of code dedicated to creating and saving an Excel workbook for your records.

                //Write the new record to the Excel sheet.
                string month = DateTime.Now.ToString("MMMM");
                int cYear = DateTime.Now.Year;

                //Write the new record to the Excel sheet.
                string path = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                string dirPath = path + "/Check-In Library/";
                DirectoryInfo di = Directory.CreateDirectory(dirPath);
                string fullPath = dirPath + month + " " + cYear + " Check-In.xlsx";
                FileInfo excelFile = new FileInfo(fullPath);
                using (ExcelPackage excel = new ExcelPackage(excelFile))
                {

                    ExcelWorksheet worksheet = excel.Workbook.Worksheets.FirstOrDefault(x => x.Name == "Worksheet1");
                    List<string[]> headerRow = new List<string[]>()
                            {
                                new string[] { "Model #", "Serial #", "Year of Manufacture", "Walk-In or Takeback", "Pressurized?", "Valve",
                                    "O-Ring 1", "O-Ring 2","Size", "Chemical", "Extra Parts", "Extra Labor", "Collar", "CT Label", "Customer","Town",
                                    "Date Arrived", "Claim Tag #", "Extg. # in Order", "Order Total","Current Year","Tally","Work Needed", "RC or Recondition"}
                            };

                    string headerRange = "A1:" + Char.ConvertFromUtf32(headerRow[0].Length + 64) + "1";
                    int colCount = 15;
                    bool emptyRow = false;
                    int rowCounter = 1;
                    bool emptyCell = false;
                    int recCounter = 1;
                    int total = 0;
                    while (emptyRow == false)
                    {

                        if (worksheet.Cells[rowCounter, colCount].Value == null)
                        {

                            emptyRowAddress = rowCounter;

                            emptyRow = true;
                        }
                        else
                        {
                            rowCounter++;
                        }

                    }
                    while (emptyCell == false)
                    {

                        if (worksheet.Cells[recCounter, 1].Value == null)
                        {

                            emptyRecRowAddress = recCounter;

                            emptyCell = true;
                        }
                        else
                        {
                            recCounter++;
                        }
                    }

                    int loopCount = emptyRowAddress;
                    int loopRowCount = 1;
                    var CustRecords = new[]
                    {
                        new
                        {
                            cust = custNameData,
                            date = dateRecData,
                            claim = claimNumData,
                            town = custTownData,
                        }
                    };
                    if (emptyRowAddress == 2)
                    {
                        foreach (var data in CustRecords)
                        {
                            worksheet.Cells[emptyRowAddress, 15].Value = data.cust;
                            worksheet.Cells[emptyRowAddress, 16].Value = data.town;
                            worksheet.Cells[emptyRowAddress, 17].Value = data.date;
                            worksheet.Cells[emptyRowAddress, 18].Value = data.claim;
                            worksheet.Cells[headerRange].AutoFitColumns();
                        }
                        // This whole section below is made to retroactively update all existing customer records.
                        int totalLoop = 2;
                        int emptyRowNum = 0;
                        bool emptyRowBool = false;
                        // This function is designed to find the first empty row in the customer data section.
                        while (emptyRowBool == false)
                        {
                            if (worksheet.Cells[totalLoop, 15].Value == null)
                            {

                                emptyRowNum = totalLoop;

                                emptyRowBool = true;
                            }
                            else
                            {
                                totalLoop++;
                            }
                        }

                        totalLoop = 2;
                        // This first 'while' loop is made to loop through the entirety of all inputted data in the customer data section.
                        // Specifically, it does this once for every customer in the list.
                        while (totalLoop < emptyRowNum)
                        {
                            string iterCustName;
                            string iterCustTown;
                            string iterCustDate;
                            string refCustName;
                            string refCustTown;
                            string refCustDate;
                            int iterRowNum = 2;
                            int matchAddress = 0;
                            total = 0;
                            refCustName = worksheet.Cells[totalLoop, 15].Value.ToString();
                            refCustTown = worksheet.Cells[totalLoop, 16].Value.ToString();
                            refCustDate = worksheet.Cells[totalLoop, 17].Value.ToString();
                            while (iterRowNum < emptyRowNum)
                            {
                                iterCustName = worksheet.Cells[iterRowNum, 15].Value.ToString();
                                iterCustTown = worksheet.Cells[iterRowNum, 16].Value.ToString();
                                iterCustDate = worksheet.Cells[iterRowNum, 17].Value.ToString();
                                if ((refCustName == iterCustName) && (refCustTown == iterCustTown) && (refCustDate == iterCustDate))
                                {
                                    total++;
                                    matchAddress = iterRowNum;
                                    worksheet.Cells[matchAddress, 19].Value = total;
                                    iterRowNum++;
                                }
                                else
                                {
                                    iterRowNum++;
                                }
                            }
                            worksheet.Cells[totalLoop, 20].Value = total;
                            totalLoop++;
                        }

                        totalLoop = 2;
                        // This first 'while' loop is made to loop through the entirety of all inputted data in the customer data section.
                        // Specifically, it does this once for every customer in the list.
                        while (totalLoop < emptyRowNum)
                        {
                            string iterCustName;
                            string iterCustTown;
                            string iterCustDate;
                            string refCustName;
                            string refCustTown;
                            string refCustDate;
                            int iterRowNum = 2;
                            total = 0;
                            refCustName = worksheet.Cells[totalLoop, 15].Value.ToString();
                            refCustTown = worksheet.Cells[totalLoop, 16].Value.ToString();
                            refCustDate = worksheet.Cells[totalLoop, 17].Value.ToString();
                            while (iterRowNum < emptyRowNum)
                            {
                                iterCustName = worksheet.Cells[iterRowNum, 15].Value.ToString();
                                iterCustTown = worksheet.Cells[iterRowNum, 16].Value.ToString();
                                iterCustDate = worksheet.Cells[iterRowNum, 17].Value.ToString();
                                if ((refCustName == iterCustName) && (refCustTown == iterCustTown) && (refCustDate == iterCustDate))
                                {
                                    total++;
                                    iterRowNum++;
                                }
                                else
                                {
                                    iterRowNum++;
                                }
                            }
                            worksheet.Cells[totalLoop, 20].Value = total;
                            string decRowTally = worksheet.Cells[totalLoop, 19].Value.ToString();
                            string decRowTotal = worksheet.Cells[totalLoop, 20].Value.ToString();
                            string cell21 = decRowTally + " of " + decRowTotal;
                            worksheet.Cells[totalLoop, 22].Value = cell21;
                            worksheet.Cells[headerRange].AutoFitColumns();
                            totalLoop++;
                        }
                        excel.SaveAs(excelFile);
                        loopCount++;
                        loopRowCount++;
                    }
                    else
                    {
                        foreach (var data in CustRecords)
                        {
                            worksheet.Cells[emptyRowAddress, 15].Value = data.cust;
                            worksheet.Cells[emptyRowAddress, 16].Value = data.town;
                            worksheet.Cells[emptyRowAddress, 17].Value = data.date;
                            worksheet.Cells[emptyRowAddress, 18].Value = data.claim;
                            worksheet.Cells[headerRange].AutoFitColumns();

                        }
                        // This whole section below is made to retroactively update all existing customer records.
                        int totalLoop = 2;
                        int emptyRowNum = 0;
                        bool emptyRowBool = false;
                        // This function is designed to find the first empty row in the customer data section.
                        while (emptyRowBool == false)
                        {
                            if (worksheet.Cells[totalLoop, 15].Value == null)
                            {

                                emptyRowNum = totalLoop;

                                emptyRowBool = true;
                            }
                            else
                            {
                                totalLoop++;
                            }
                        }

                        totalLoop = 2;
                        // This first 'while' loop is made to loop through the entirety of all inputted data in the customer data section.
                        // Specifically, it does this once for every customer in the list.
                        while (totalLoop < emptyRowNum)
                        {
                            string iterCustName;
                            string iterCustTown;
                            string iterCustDate;
                            string refCustName;
                            string refCustTown;
                            string refCustDate;
                            int iterRowNum = 2;
                            int matchAddress = 0;
                            total = 0;
                            refCustName = worksheet.Cells[totalLoop, 15].Value.ToString();
                            refCustTown = worksheet.Cells[totalLoop, 16].Value.ToString();
                            refCustDate = worksheet.Cells[totalLoop, 17].Value.ToString();
                            while (iterRowNum < emptyRowNum)
                            {
                                iterCustName = worksheet.Cells[iterRowNum, 15].Value.ToString();
                                iterCustTown = worksheet.Cells[iterRowNum, 16].Value.ToString();
                                iterCustDate = worksheet.Cells[iterRowNum, 17].Value.ToString();
                                if ((refCustName == iterCustName) && (refCustTown == iterCustTown) && (refCustDate == iterCustDate))
                                {
                                    total++;
                                    matchAddress = iterRowNum;
                                    worksheet.Cells[matchAddress, 19].Value = total;
                                    iterRowNum++;
                                }
                                else
                                {
                                    iterRowNum++;
                                }
                            }
                            worksheet.Cells[totalLoop, 20].Value = total;
                            totalLoop++;
                        }

                        totalLoop = 2;
                        // This first 'while' loop is made to loop through the entirety of all inputted data in the customer data section.
                        // Specifically, it does this once for every customer in the list.
                        while (totalLoop < emptyRowNum)
                        {
                            string iterCustName;
                            string iterCustTown;
                            string iterCustDate;
                            string refCustName;
                            string refCustTown;
                            string refCustDate;
                            int iterRowNum = 2;
                            total = 0;
                            refCustName = worksheet.Cells[totalLoop, 15].Value.ToString();
                            refCustTown = worksheet.Cells[totalLoop, 16].Value.ToString();
                            refCustDate = worksheet.Cells[totalLoop, 17].Value.ToString();
                            while (iterRowNum < emptyRowNum)
                            {
                                iterCustName = worksheet.Cells[iterRowNum, 15].Value.ToString();
                                iterCustTown = worksheet.Cells[iterRowNum, 16].Value.ToString();
                                iterCustDate = worksheet.Cells[iterRowNum, 17].Value.ToString();
                                if ((refCustName == iterCustName) && (refCustTown == iterCustTown) && (refCustDate == iterCustDate))
                                {
                                    total++;
                                    iterRowNum++;
                                }
                                else
                                {
                                    iterRowNum++;
                                }
                            }
                            worksheet.Cells[totalLoop, 20].Value = total;
                            string decRowTally = worksheet.Cells[totalLoop, 19].Value.ToString();
                            string decRowTotal = worksheet.Cells[totalLoop, 20].Value.ToString();
                            string cell21 = decRowTally + " of " + decRowTotal;
                            worksheet.Cells[totalLoop, 22].Value = cell21;
                            worksheet.Cells[headerRange].AutoFitColumns();
                            totalLoop++;
                        }
                        excel.SaveAs(excelFile);
                        loopCount++;
                        loopRowCount++;
                        if (worksheet.Cells[emptyRowAddress, 1].Value == null)
                        {
                            SystemSounds.Hand.Play();
                            MessageBox.Show("Customer Info has surpassed Extg Info", "OVERLOAD", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }


                    }
                    e.Handled = true;
                    e.SuppressKeyPress = true;
                    if (sameCust == false)
                    {
                        Control_ReturnFocus();
                    }
                    else
                    {
                        SameCustomerFocus();
                    }
                    ExcelReadoutRefresh();
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
            this.custInfoSubmit = new System.Windows.Forms.Button();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.sameCustomer = new System.Windows.Forms.CheckBox();
            this.custTown = new System.Windows.Forms.TextBox();
            this.townLabel = new System.Windows.Forms.Label();
            this.claimNum = new System.Windows.Forms.TextBox();
            this.dateRec = new System.Windows.Forms.TextBox();
            this.custName = new System.Windows.Forms.TextBox();
            this.claimLabel = new System.Windows.Forms.Label();
            this.dateRecLabel = new System.Windows.Forms.Label();
            this.custNameLabel = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.panel1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.SuspendLayout();
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.custInfoSubmit);
            this.panel1.Controls.Add(this.dataGridView1);
            this.panel1.Controls.Add(this.sameCustomer);
            this.panel1.Controls.Add(this.custTown);
            this.panel1.Controls.Add(this.townLabel);
            this.panel1.Controls.Add(this.claimNum);
            this.panel1.Controls.Add(this.dateRec);
            this.panel1.Controls.Add(this.custName);
            this.panel1.Controls.Add(this.claimLabel);
            this.panel1.Controls.Add(this.dateRecLabel);
            this.panel1.Controls.Add(this.custNameLabel);
            this.panel1.Controls.Add(this.label1);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel1.Location = new System.Drawing.Point(0, 0);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(800, 450);
            this.panel1.TabIndex = 0;
            // 
            // custInfoSubmit
            // 
            this.custInfoSubmit.Anchor = System.Windows.Forms.AnchorStyles.Top;
            this.custInfoSubmit.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F);
            this.custInfoSubmit.Location = new System.Drawing.Point(342, 203);
            this.custInfoSubmit.Name = "custInfoSubmit";
            this.custInfoSubmit.Size = new System.Drawing.Size(96, 31);
            this.custInfoSubmit.TabIndex = 10;
            this.custInfoSubmit.Text = "&Submit";
            this.custInfoSubmit.UseVisualStyleBackColor = true;
            this.custInfoSubmit.Click += new System.EventHandler(this.custInfoSubmit_Click);
            // 
            // dataGridView1
            // 
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.dataGridView1.Location = new System.Drawing.Point(0, 240);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.Size = new System.Drawing.Size(800, 210);
            this.dataGridView1.TabIndex = 9;
            // 
            // sameCustomer
            // 
            this.sameCustomer.Anchor = System.Windows.Forms.AnchorStyles.Top;
            this.sameCustomer.AutoSize = true;
            this.sameCustomer.Location = new System.Drawing.Point(332, 89);
            this.sameCustomer.Name = "sameCustomer";
            this.sameCustomer.Size = new System.Drawing.Size(106, 17);
            this.sameCustomer.TabIndex = 8;
            this.sameCustomer.Text = "Same Customer?";
            this.sameCustomer.UseVisualStyleBackColor = true;
            this.sameCustomer.CheckedChanged += new System.EventHandler(this.sameCustomer_Checked);
            // 
            // custTown
            // 
            this.custTown.Anchor = System.Windows.Forms.AnchorStyles.Top;
            this.custTown.Location = new System.Drawing.Point(249, 167);
            this.custTown.Name = "custTown";
            this.custTown.Size = new System.Drawing.Size(100, 20);
            this.custTown.TabIndex = 5;
            this.custTown.TextChanged += new System.EventHandler(this.custTown_TextChanged);
            this.custTown.Enter += new System.EventHandler(this.custTown_Enter);
            this.custTown.KeyDown += new System.Windows.Forms.KeyEventHandler(this.Control_KeyUp);
            // 
            // townLabel
            // 
            this.townLabel.Anchor = System.Windows.Forms.AnchorStyles.Top;
            this.townLabel.AutoSize = true;
            this.townLabel.Location = new System.Drawing.Point(256, 131);
            this.townLabel.Name = "townLabel";
            this.townLabel.Size = new System.Drawing.Size(81, 13);
            this.townLabel.TabIndex = 7;
            this.townLabel.Text = "Customer Town";
            // 
            // claimNum
            // 
            this.claimNum.Anchor = System.Windows.Forms.AnchorStyles.Top;
            this.claimNum.Location = new System.Drawing.Point(606, 167);
            this.claimNum.Name = "claimNum";
            this.claimNum.Size = new System.Drawing.Size(100, 20);
            this.claimNum.TabIndex = 7;
            this.claimNum.TextChanged += new System.EventHandler(this.claimNum_TextChanged);
            this.claimNum.Enter += new System.EventHandler(this.claimNum_Enter);
            this.claimNum.KeyDown += new System.Windows.Forms.KeyEventHandler(this.Control_SaveRecord);
            // 
            // dateRec
            // 
            this.dateRec.Anchor = System.Windows.Forms.AnchorStyles.Top;
            this.dateRec.Location = new System.Drawing.Point(431, 167);
            this.dateRec.Name = "dateRec";
            this.dateRec.Size = new System.Drawing.Size(100, 20);
            this.dateRec.TabIndex = 6;
            this.dateRec.TextChanged += new System.EventHandler(this.dateRec_TextChanged);
            this.dateRec.Enter += new System.EventHandler(this.dateRec_Enter);
            this.dateRec.KeyDown += new System.Windows.Forms.KeyEventHandler(this.Control_KeyUp);
            // 
            // custName
            // 
            this.custName.Anchor = System.Windows.Forms.AnchorStyles.Top;
            this.custName.Location = new System.Drawing.Point(64, 167);
            this.custName.Name = "custName";
            this.custName.Size = new System.Drawing.Size(100, 20);
            this.custName.TabIndex = 4;
            this.custName.TextChanged += new System.EventHandler(this.custName_TextChanged);
            this.custName.Enter += new System.EventHandler(this.custName_Enter);
            this.custName.KeyDown += new System.Windows.Forms.KeyEventHandler(this.Control_KeyUp);
            // 
            // claimLabel
            // 
            this.claimLabel.Anchor = System.Windows.Forms.AnchorStyles.Top;
            this.claimLabel.AutoSize = true;
            this.claimLabel.Location = new System.Drawing.Point(619, 130);
            this.claimLabel.Name = "claimLabel";
            this.claimLabel.Size = new System.Drawing.Size(64, 13);
            this.claimLabel.TabIndex = 3;
            this.claimLabel.Text = "Claim Tag #";
            // 
            // dateRecLabel
            // 
            this.dateRecLabel.Anchor = System.Windows.Forms.AnchorStyles.Top;
            this.dateRecLabel.AutoSize = true;
            this.dateRecLabel.Location = new System.Drawing.Point(440, 130);
            this.dateRecLabel.Name = "dateRecLabel";
            this.dateRecLabel.Size = new System.Drawing.Size(79, 13);
            this.dateRecLabel.TabIndex = 2;
            this.dateRecLabel.Text = "Date Received";
            // 
            // custNameLabel
            // 
            this.custNameLabel.Anchor = System.Windows.Forms.AnchorStyles.Top;
            this.custNameLabel.AutoSize = true;
            this.custNameLabel.Location = new System.Drawing.Point(72, 131);
            this.custNameLabel.Name = "custNameLabel";
            this.custNameLabel.Size = new System.Drawing.Size(82, 13);
            this.custNameLabel.TabIndex = 1;
            this.custNameLabel.Text = "Customer Name";
            // 
            // label1
            // 
            this.label1.Anchor = System.Windows.Forms.AnchorStyles.Top;
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 20.25F, System.Drawing.FontStyle.Bold);
            this.label1.Location = new System.Drawing.Point(214, 26);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(374, 31);
            this.label1.TabIndex = 0;
            this.label1.Text = "Enter customer information.";
            // 
            // CustomerInfo
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.panel1);
            this.Name = "CustomerInfo";
            this.Text = "Customer Info";
            this.Load += new System.EventHandler(this.CustomerInfo_Load);
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox claimNum;
        private System.Windows.Forms.TextBox dateRec;
        private System.Windows.Forms.TextBox custName;
        private System.Windows.Forms.Label claimLabel;
        private System.Windows.Forms.Label dateRecLabel;
        private System.Windows.Forms.Label custNameLabel;
        private Label townLabel;
        private TextBox custTown;
        private CheckBox sameCustomer;
        private DataGridView dataGridView1;
        private Button custInfoSubmit;
    }
}