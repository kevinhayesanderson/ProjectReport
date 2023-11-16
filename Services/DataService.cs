// Ignore Spelling: Tohhmm

using Models;
using Utilities;

namespace Services
{
    public class DataService(ILogger logger)
    {
        public long[] Months => [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12];

        public List<ConsolidatedData> Consolidate(PtrData ptrData, MonthlyReportData monthlyReportData)
        {
            List<ConsolidatedData> consolidatedDataList = [];
            try
            {
                logger.LogInfo("Consolidating data", 2);
                IEnumerable<string> projectIds = ptrData.ProjectIds.Union(monthlyReportData.ProjectIds);
                consolidatedDataList.AddRange(from projectId in projectIds
                                              let consolidatedData = new ConsolidatedData
                                              {
                                                  ProjectId = projectId,
                                                  TotalEffort = ptrData.ProjectEfforts.TryGetValue(projectId, out TimeSpan value) ? value : TimeSpan.Zero,
                                                  EmployeeActualEffort = monthlyReportData.EmployeesData
                                                                        .Where(ed => ed.ProjectData.ContainsKey(projectId))
                                                                        .Select(ed => new EmployeeActualEffort()
                                                                        {
                                                                            Id = ed.Id,
                                                                            Name = ed.Name,
                                                                            ProjectId = projectId,
                                                                            ActualEffort = ed.ProjectData[projectId]
                                                                        }).ToList()
                                              }
                                              select consolidatedData);
            }
            catch (Exception ex)
            {
                logger.LogErrorAndExit("Error on consolidating data: " + ex.ToString() + " ");
            }
            return consolidatedDataList;
        }

        public List<string> GetFyMonths(string financialYear)
        {
            string[] strArray = financialYear.Split('-');
            return
            [
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
            ];
        }

        public string TohhmmFormatString(TimeSpan timeSpan)
        {
            string totalHours = ((int)timeSpan.TotalHours).ToString();
            string minutes = timeSpan.Minutes.ToString();
            if (timeSpan.Seconds == 59 && (timeSpan.Milliseconds == 999 || timeSpan.Microseconds == 999))
            {
                if (timeSpan.Minutes + 1 == 60)
                {
                    minutes = "00";
                    totalHours = ((int)timeSpan.TotalHours + 1).ToString();
                }
                else
                {
                    minutes = (timeSpan.Minutes + 1).ToString();
                }
            }
            return $"{totalHours}:{minutes}";
        }
    }
}