using Microsoft.Office.Interop.Excel;
using Models;
using Utilities;

using Excel = Microsoft.Office.Interop.Excel;

namespace Services
{
    public class WriteService(ILogger logger)
    {
        public bool WriteInOutEntry(List<string> monthlyReports, MusterOptionsDatas musterOptionsDatas)
        {
            bool res = true;
            logger.LogInfo("Writing InOutEntry in monthly reports:", 1);
            Application excelApp = new()
            {
                Visible = true // Optional, make Excel visible
            };
            try
            {
                IEnumerable<Workbook> workbooks = GetMonthlyReportsWorkbooks(excelApp, monthlyReports);

                foreach ((uint employeeId, MusterOptionsData musterOptionsData) in musterOptionsDatas.Datas)
                {
                }

                SaveAndCloseWorkbooks(workbooks);

                // Open the existing workbook with password authentication
                Workbook workbook = excelApp.Workbooks.Open("", Password: "");

                // Get the first worksheet (you can change the index if you have multiple sheets)
                Worksheet worksheet = (Worksheet)workbook.Sheets[1];

                // Example: Write data to cell A1
                WriteDataToCell(worksheet, 1, 1, "Hello, Excel!");

                // Save the changes
                workbook.Save();

                // Close the workbook
                workbook.Close();
            }
            catch (Exception ex)
            {
                logger.LogError($"An error occurred on writing InOutEntry in monthly reports: {ex.Message}");
                return false;
            }
            finally
            {
                excelApp.Quit();
            }
            return res;
        }

        private void SaveAndCloseWorkbooks(IEnumerable<Workbook> workbooks)
        {
            foreach (var workbook in workbooks)
            {
                workbook.Save();
                workbook.Close();
            }
        }

        private IEnumerable<Workbook> GetMonthlyReportsWorkbooks(Application excelApp, List<string> monthlyReports)
        {
            return monthlyReports.Select(mr => excelApp.Workbooks.Open(mr));
        }

        // Helper method to write data to a specific cell
        private static void WriteDataToCell(Worksheet worksheet, int row, int column, object data)
        {
            Excel.Range cell = (Excel.Range)worksheet.Cells[row, column];
            cell.Value = data;
        }
    }

    ////public class ExcelFile
    ////{
    ////    private string excelFilePath = string.Empty;
    ////    private int rowNumber = 1; // define first row number to enter data in excel

    ////    private Excel.Application myExcelApplication;
    ////    private Excel.Workbook myExcelWorkbook;
    ////    private Excel.Worksheet myExcelWorkSheet;

    ////    public string ExcelFilePath
    ////    {
    ////        get { return excelFilePath; }
    ////        set { excelFilePath = value; }
    ////    }

    ////    public int Rownumber
    ////    {
    ////        get { return rowNumber; }
    ////        set { rowNumber = value; }
    ////    }

    ////    public void openExcel()
    ////    {
    ////        myExcelApplication = null;

    ////        myExcelApplication = new Excel.Application
    ////        {
    ////            DisplayAlerts = false // turn off alerts
    ////        }; // create Excell App

    ////        myExcelWorkbook = myExcelApplication.Workbooks._Open(excelFilePath, Missing.Value,
    ////           Missing.Value, Missing.Value, Missing.Value,
    ////           Missing.Value, Missing.Value, Missing.Value,
    ////           Missing.Value, Missing.Value, Missing.Value,
    ////           Missing.Value, Missing.Value); // open the existing excel file

    ////        int numberOfWorkbooks = myExcelApplication.Workbooks.Count; // get number of workbooks (optional)

    ////        myExcelWorkSheet = (Excel.Worksheet)myExcelWorkbook.Worksheets[1]; // define in which worksheet, do you want to add data
    ////        myExcelWorkSheet.Name = "WorkSheet 1"; // define a name for the worksheet (optinal)

    ////        int numberOfSheets = myExcelWorkbook.Worksheets.Count; // get number of worksheets (optional)
    ////    }

    ////    public void addDataToExcel(string firstname, string lastname, string language, string email, string company)
    ////    {
    ////        myExcelWorkSheet.Cells[rowNumber, "H"] = firstname;
    ////        myExcelWorkSheet.Cells[rowNumber, "J"] = lastname;
    ////        myExcelWorkSheet.Cells[rowNumber, "Q"] = language;
    ////        myExcelWorkSheet.Cells[rowNumber, "BH"] = email;
    ////        myExcelWorkSheet.Cells[rowNumber, "CH"] = company;
    ////        rowNumber++;  // if you put this method inside a loop, you should increase rownumber by one or wat ever is your logic
    ////    }

    ////    public void closeExcel()
    ////    {
    ////        try
    ////        {
    ////            myExcelWorkbook.SaveAs(excelFilePath, Missing.Value, Missing.Value, Missing.Value,
    ////                                           Missing.Value, Missing.Value, Excel.XlSaveAsAccessMode.xlNoChange,
    ////                                           Missing.Value, Missing.Value, Missing.Value,
    ////                                           Missing.Value, Missing.Value); // Save data in excel

    ////            myExcelWorkbook.Close(true, excelFilePath, Missing.Value); // close the worksheet
    ////        }
    ////        finally
    ////        {
    ////            if (myExcelApplication != null)
    ////            {
    ////                myExcelApplication.Quit(); // close the excel application
    ////            }
    ////        }
    ////    }
    ////}
}