using Models;
using Utilities;

namespace Services
{
    public static class DataService
    {
        public static List<ConsolidatedData> Consolidate(
          PtrData ptrData,
          MonthlyReportData monthlyReportData)
        {
            List<ConsolidatedData> consolidatedDataList = new List<ConsolidatedData>();
            try
            {
                ConsoleLogger.LogInfo("Consolidating data.", 1);
                consolidatedDataList = ptrData.ProjectEfforts.Select(projEffort => new ConsolidatedData()
                {
                    ProjectId = projEffort.Key,
                    TotalEffort = projEffort.Value,
                    EmployeeActualEffort = monthlyReportData.EmployeesData.Where(ed => ed.ProjectTime.ContainsKey(projEffort.Key.Trim())).Select(ed => new EmployeeActualEffort()
                    {
                        Id = ed.Id,
                        Name = ed.Name,
                        ProjectId = projEffort.Key,
                        ActualEffort = ed.ProjectTime[projEffort.Key]
                    }).ToList()
                }).ToList();
            }
            catch (Exception ex)
            {
                ConsoleLogger.LogErrorAndExit("Error on consolidating data: " + ex.Message + " ");
            }
            return consolidatedDataList;
        }

        public static List<string> GetFyMonths(string financialYear)
        {
            string[] strArray = financialYear.Split('-');
            return new List<string>()
      {
        "Apr-" + strArray[0],
        "May-" + strArray[0],
        "Jun-" + strArray[0],
        "Jul-" + strArray[0],
        "Aug-" + strArray[0],
        "Sep-" + strArray[0],
        "Oct-" + strArray[0],
        "Nov-" + strArray[0],
        "Dec-" + strArray[0],
        "Jan-" + strArray[1],
        "Feb-" + strArray[1],
        "Mar-" + strArray[1]
      };
        }
    }
}
