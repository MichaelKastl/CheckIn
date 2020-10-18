using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.IO;
using System.Linq;

namespace CheckInGUI
{

    class MegaBaseBuilder
    {
        public static int rowCnt;
        internal static bool SheetExist(string fullFilePath, string sheetName)
        {
            using (var package = new ExcelPackage(new FileInfo(fullFilePath)))
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                return package.Workbook.Worksheets.Any(sheet => sheet.Name == sheetName);
            }
        }
        public static void MegaBaseBuilderMethod(ExcelWorksheet workSheet, string name)
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
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                // Create the worksheet, if it doesn't exist already.
                if (SheetExist(megaPath, "Serial Num. Database") == false)
                {
                    megaPkg.Workbook.Worksheets.Add("Serial Num. Database");
                }

                // Initialize worksheet, add and format header row.
                ExcelWorksheet megaWorksheet = megaPkg.Workbook.Worksheets[0];
                megaWorksheet.Cells["A1"].Value = "Customer Name";
                megaWorksheet.Cells["B1"].Value = "Customer Town";
                megaWorksheet.Cells["C1"].Value = "Model Num.";
                megaWorksheet.Cells["D1"].Value = "Serial Num.";
                megaWorksheet.Cells["A1:D1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                megaWorksheet.Cells["A1:D1"].Style.Font.Size = 14;
                megaWorksheet.Cells["A1:D1"].Style.Font.Bold = true;
                megaWorksheet.Cells["A1:D1"].AutoFitColumns();

                // Start by trimming the name variable to just the Customer Name.
                int rowStart = workSheet.Dimension.Start.Row;
                int rowEnd = workSheet.Dimension.End.Row + 1;

                string cellRange = rowStart.ToString() + ":" + rowEnd.ToString();
                int index = name.LastIndexOf("-");
                string custName = name.Substring(0, index - 1);

                //Next, trim to just the Customer Town.
                int townLength = name.Length - 2 - custName.Length;
                string custTown = name.Substring(index + 1, townLength);

                // Now, search for the first occurrence of that Customer Name in the document, as well as how many times it repeats.
                var searchCell = from cell in workSheet.Cells[cellRange]
                                 where cell.Value != null && cell.Value.ToString() == custName && cell.Start != null
                                 select cell.Start.Row;

                int rowNum = searchCell.FirstOrDefault();
                int nameCount = searchCell.Count();

                // Loop for as many times as the Customer Name appears.
                while (nameCount > 0)
                {
                    rowCnt = megaWorksheet.Dimension.End.Row + 1;
                    // Set modelNum and serialNum to the values of the cells in the rows where Customer Name was found.
                    string modelNum = workSheet.Cells["A" + rowNum].Value.ToString();
                    string serialNum = workSheet.Cells["B" + rowNum].Value.ToString();

                    megaWorksheet.Cells["A" + rowCnt].Value = custName;
                    megaWorksheet.Cells["B" + rowCnt].Value = custTown;
                    megaWorksheet.Cells["C" + rowCnt].Value = modelNum;
                    megaWorksheet.Cells["D" + rowCnt].Value = serialNum;

                    rowCnt++;

                    // Set the range to 1 row past the previous occurence of Customer Name.
                    rowNum++;

                    string cellRange2 = rowNum.ToString() + ":" + rowEnd.ToString();
                    var searchCell2 = from cell in workSheet.Cells[cellRange2]
                                      where cell.Value != null && cell.Value.ToString() == custName && cell.Start != null
                                      select cell.Start.Row;

                    // Find the next occurence of Customer Name.
                    rowNum = searchCell2.FirstOrDefault();

                    // Remove duplicate values, and save file.
                    megaPkg.SaveAs(megaFile);

                    // Decrement the loop counter, and go again.
                    Console.WriteLine("Record Entered: " + custName + " " + custTown + " " + modelNum + " " + serialNum + ".");
                    nameCount--;
                }
            }
        }
    }
}
