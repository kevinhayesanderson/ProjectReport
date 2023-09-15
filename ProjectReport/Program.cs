using Microsoft.Extensions.FileSystemGlobbing;
using Models;
using Services;
using System.Globalization;
using System.Text;
using Utilities;

namespace ConsoleApplication
{
    internal static class Program
    {
        private const string ApplicationDev = "kevin.hayes@ambigai.net";
        private static string _exportFolder = string.Empty;
        private static string _time = string.Empty;

        private static List<ConsolidatedData>? ConsolidatedData { get; set; }

        private static MonthlyReportData? MonthlyReportData { get; set; }

        private static List<string>? MonthlyReports { get; set; }

        private static PtrData? PtrData { get; set; }

        private static List<string>? PtrFiles { get; set; }

        private static UserSettings? UserSettings { get; set; }

        private static void Main()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            ConsoleLogger.LogInfo("Running Project Report Application.");
            UserSettings = ReadService.ReadUserSettings();
            if (Directory.Exists(UserSettings.Folder))
            {
                _time = DateTime.Now.ToString("t", CultureInfo.InvariantCulture).Replace(":", "_");
                _exportFolder = UserSettings.Folder + "\\Reports_" + _time;
                Matcher matcher1 = new();
                matcher1.AddInclude("*Monthly_Report*");
                MonthlyReports = matcher1.GetResultsInFullPath(UserSettings.Folder).ToList();
                if (MonthlyReports.Count > 0)
                {
                    ConsoleLogger.LogInfo("Monthly reports found:", 1);
                    MonthlyReports.ForEach(mr => ConsoleLogger.Log(new FileInfo(mr).Name));
                    if (UserSettings.GenerateLeaveReport)
                    {
                        ExportService.ExportLeaveReport(MonthlyReports, UserSettings.FinancialYear, _time, _exportFolder);
                        ConsoleLogger.ExitApplication();
                    }
                    else
                    {
                        MonthlyReportData = ReadService.ReadMonthlyReports(MonthlyReports, UserSettings);
                        ExportService.ExportMonthlyReportInter(MonthlyReportData, _time, _exportFolder);
                    }
                }
                else
                    ConsoleLogger.LogWarningAndExit("No Monthly reports found on " + UserSettings.Folder + ", needed monthly reports to generate consolidated report or leave report.");
                Matcher matcher2 = new();
                matcher2.AddInclude("*ACS_PTR*");
                PtrFiles = matcher2.GetResultsInFullPath(UserSettings.Folder).ToList();
                if (PtrFiles.Count > 0)
                {
                    ConsoleLogger.LogInfo("PTR's found:", 1);
                    PtrFiles.ForEach(ptr => ConsoleLogger.Log(new FileInfo(ptr).Name));
                    PtrData = ReadService.ReadPtr(PtrFiles, UserSettings);
                    ExportService.ExportPtrInter(PtrData, _time, _exportFolder);
                }
                else
                    ConsoleLogger.LogWarningAndExit("No PTR found on " + UserSettings.Folder + ", needed PTR to generate consolidated report.");
                if (MonthlyReportData != null && PtrData != null)
                {
                    ConsolidatedData = DataService.Consolidate(PtrData, MonthlyReportData);
                    if (ConsolidatedData.Count > 0)
                        ExportService.ExportConsolidateData(ConsolidatedData, PtrData, MonthlyReportData, _time, _exportFolder);
                    else
                        ConsoleLogger.LogWarningAndExit($"Consolidated data is empty, modify filter conditions or check if data is present, otherwise report application error to {ApplicationDev}", 2);
                }
                else
                    ConsoleLogger.LogWarningAndExit($"Either PTR or Monthly report data is empty, modify filter conditions or check if data is present, otherwise report application error to {ApplicationDev}", 2);
            }
            else
                ConsoleLogger.LogErrorAndExit("Directory doesn't exist: " + UserSettings.Folder, 2);
            ConsoleLogger.ExitApplication();
        }
    }
}