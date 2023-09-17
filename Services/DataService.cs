﻿using Models;
using Utilities;

namespace Services
{
    public static class DataService
    {
        public static readonly int[] Months = new int[12] { 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12 };

        public static List<ConsolidatedData> Consolidate(PtrData ptrData, MonthlyReportData monthlyReportData)
        {
            List<ConsolidatedData> consolidatedDataList = new();
            try
            {
                ConsoleLogger.LogInfo("Consolidating data", 2);
                var projectIds = ptrData.ProjectIds.Union(monthlyReportData.ProjectIds);
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