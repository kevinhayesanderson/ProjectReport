using Microsoft.Office.Interop.Excel;
using Models;
using Utilities;

namespace Services
{
    public class WriteService(ILogger logger)
    {
        public bool WriteInOutEntry(List<string> monthlyReports, MusterOptionsDatas musterOptionsDatas)
        {
            bool res = true;
            logger.LogInfo("Writing InOutEntry in monthly reports:", 1);
            Application excelApp = new();
            try
            {
                IEnumerable<Workbook> workbooks = GetMonthlyReportsWorkbooks(excelApp, monthlyReports);

                foreach ((uint employeeId, MusterOptionsData musterOptionsData) in musterOptionsDatas.Datas)
                {
                    //// TODO: Write musteroptions in and out time on Monthly reports.
                }

                SaveAndCloseWorkbooks(workbooks);
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

        private static IEnumerable<Workbook> GetMonthlyReportsWorkbooks(Application excelApp, List<string> monthlyReports)
        {
            return monthlyReports.Select(mr => excelApp.Workbooks.Open(mr));
        }

        private static void SaveAndCloseWorkbooks(IEnumerable<Workbook> workbooks)
        {
            foreach (var workbook in workbooks)
            {
                workbook.Save();
                workbook.Close();
            }
        }
    }
}