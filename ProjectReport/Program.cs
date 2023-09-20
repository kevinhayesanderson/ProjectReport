using Microsoft.Extensions.FileSystemGlobbing;
using Models;
using Services;
using System.Diagnostics;
using System.Text;
using System.Text.RegularExpressions;
using Utilities;

namespace ProjectReport
{
    internal static partial class Program
    {
        private const string ApplicationDev = "kevin.hayes@ambigai.net";
        private static string _exportFolder = string.Empty;
        private static MonthlyReportData _monthlyReportData = new();
        private static PtrData _ptrData = new();
        private static string _time = string.Empty;
        public static ref MonthlyReportData MonthlyReportData => ref _monthlyReportData;
        public static ref PtrData PtrData => ref _ptrData;

        private static void Main()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            _time = RemoveAllSymbols().Replace(DateTime.Now.ToString(), "_");
            Console.Title = $"Project Report PID:{Process.GetCurrentProcess().Id} {_time}";
            ConsoleLogger.LogInfo($"Running Project Report Application at {_time}");
            ReadService.ReadUserSettings(out UserSettings userSettings);
            if (Directory.Exists(userSettings?.Folder))
            {
                _exportFolder = $"{userSettings.Folder}\\Reports_{_time}";
                Matcher monthlyReportMatcher = new();
                _ = monthlyReportMatcher.AddInclude("*Monthly_Report*");
                List<string> MonthlyReports = monthlyReportMatcher.GetResultsInFullPath(userSettings.Folder).ToList();
                if (MonthlyReports.Count > 0)
                {
                    ConsoleLogger.LogInfo("Monthly reports found:", 1);
                    MonthlyReports.ForEach(mr => ConsoleLogger.Log(new FileInfo(mr).Name));
                    if (userSettings.GenerateLeaveReport)
                    {
                        ExportService.ExportLeaveReport(in MonthlyReports, userSettings.FinancialYear, in _exportFolder);
                        ConsoleLogger.ExitApplication();
                    }
                    else
                    {
                        _monthlyReportData = ReadService.ReadMonthlyReports(MonthlyReports, userSettings);
                        ExportService.ExportMonthlyReportInter(ref MonthlyReportData, in _time, in _exportFolder);
                    }
                }
                else
                {
                    ConsoleLogger.LogWarningAndExit("No Monthly reports found on " + userSettings.Folder + ", needed monthly reports to generate consolidated report or leave report.");
                }

                Matcher ptrMatcher = new();
                _ = ptrMatcher.AddInclude("*ACS_PTR*");
                List<string> PtrFiles = ptrMatcher.GetResultsInFullPath(userSettings.Folder).ToList();
                if (PtrFiles.Count > 0)
                {
                    ConsoleLogger.LogInfo("PTR's found:", 2);
                    PtrFiles.ForEach(ptr => ConsoleLogger.Log(new FileInfo(ptr).Name));
                    _ptrData = ReadService.ReadPtr(PtrFiles, userSettings);
                    ExportService.ExportPtrInter(ref PtrData, in _time, in _exportFolder);
                }
                else
                {
                    ConsoleLogger.LogWarningAndExit("No PTR found on " + userSettings.Folder + ", needed PTR to generate consolidated report.");
                }

                if (MonthlyReportData != null && PtrData != null)
                {
                    List<ConsolidatedData> ConsolidatedData = DataService.Consolidate(PtrData, MonthlyReportData);
                    if (ConsolidatedData.Count > 0)
                    {
                        ExportService.ExportConsolidateData(in ConsolidatedData, ref MonthlyReportData, in _time, in _exportFolder);
                    }
                    else
                    {
                        ConsoleLogger.LogWarningAndExit($"Consolidated data is empty, modify filter conditions or check if data is present, otherwise report application error to {ApplicationDev}", 2);
                    }
                }
                else
                {
                    ConsoleLogger.LogWarningAndExit($"Either PTR or Monthly report data is empty, modify filter conditions or check if data is present, otherwise report application error to {ApplicationDev}", 2);
                }
            }
            else
            {
                ConsoleLogger.LogErrorAndExit($"Directory doesn't exist: {userSettings?.Folder}", 2);
            }

            ConsoleLogger.ExitApplication();
        }

        [GeneratedRegex("[^\\w\\d]")]
        private static partial Regex RemoveAllSymbols();
    }
}