using Microsoft.Extensions.FileSystemGlobbing;
using Models;
using Services;
using Utilities;

namespace ProjectReport.Actions
{
    [SettingName("GenerateConsolidatedReport")]
    internal class GenerateConsolidatedReportAction : IAction
    {
        private readonly int _monthlyReportIdCol;
        private readonly object[] _monthlyReportMonths;
        private readonly int _ptrBookingMonthCol;
        private readonly object[] _ptrBookingMonths;
        private readonly object[] _ptrEffortCols;
        private readonly int _ptrProjectIdCol;
        private readonly string _ptrSheetName;
        private readonly string _time;
        private MonthlyReportData _monthlyReportData;
        private PtrData _ptrData;

        public GenerateConsolidatedReportAction(bool run, string inputFolder, string time, object[] monthlyReportMonths, int monthlyReportIdCol, int ptrBookingMonthCol, object[] ptrBookingMonths, object[] ptrEffortCols, int ptrProjectIdCol, string ptrSheetName)
        {
            (Run, InputFolder) = (run, inputFolder);
            (_time, _monthlyReportMonths, _monthlyReportIdCol, _ptrBookingMonthCol, _ptrBookingMonths, _ptrEffortCols, _ptrProjectIdCol, _ptrSheetName) =
            (time, monthlyReportMonths, monthlyReportIdCol, ptrBookingMonthCol, ptrBookingMonths, ptrEffortCols, ptrProjectIdCol, ptrSheetName);
            (_monthlyReportData, _ptrData) = (new MonthlyReportData(), new PtrData());
        }

        public string InputFolder { get; }

        public bool Run { get; }

        public bool Execute()
        {
            bool res = false;
            var _exportFolder = @$"{InputFolder}\Reports_{_time}";
            List<string> monthlyReports = Helper.GetMonthlyReports(InputFolder);
            if (monthlyReports.Count > 0)
            {
                ConsoleLogger.LogInfo("Monthly reports found:", 1);
                monthlyReports.ForEach(mr => ConsoleLogger.Log(new FileInfo(mr).Name));
                _monthlyReportData = ReadService.ReadMonthlyReports(monthlyReports, _monthlyReportMonths, _monthlyReportIdCol);
                ExportService.ExportMonthlyReportInter(ref _monthlyReportData, in _time, in _exportFolder);
            }
            else
            {
                ConsoleLogger.LogWarningAndExit($"No Monthly reports found on {InputFolder}, needed monthly reports to generate consolidated report or leave report.");
            }

            Matcher ptrMatcher = new();
            _ = ptrMatcher.AddInclude(Constants.PTRPattern);
            List<string> ptrFiles = ptrMatcher.GetResultsInFullPath(InputFolder).ToList();
            if (ptrFiles.Count > 0)
            {
                ConsoleLogger.LogInfo("PTR's found:", 2);
                ptrFiles.ForEach(ptr => ConsoleLogger.Log(new FileInfo(ptr).Name));
                _ptrData = ReadService.ReadPtr(ptrFiles, _ptrBookingMonthCol, _ptrBookingMonths, _ptrEffortCols, _ptrProjectIdCol, _ptrSheetName);
                ExportService.ExportPtrInter(ref _ptrData, in _time, in _exportFolder);
            }
            else
            {
                ConsoleLogger.LogWarningAndExit($"No PTR found on {InputFolder}, needed PTR to generate consolidated report.");
            }

            if (_monthlyReportData != null && _ptrData != null)
            {
                List<ConsolidatedData> ConsolidatedData = DataService.Consolidate(_ptrData, _monthlyReportData);
                if (ConsolidatedData.Count > 0)
                {
                    ExportService.ExportConsolidateData(in ConsolidatedData, ref _monthlyReportData, in _time, in _exportFolder);
                }
                else
                {
                    ConsoleLogger.LogWarningAndExit($"Consolidated data is empty, modify filter conditions or check if data is present, otherwise report application error to {Constants.ApplicationDev}", 2);
                }
            }
            else
            {
                ConsoleLogger.LogWarningAndExit($"Either PTR or Monthly report data is empty, modify filter conditions or check if data is present, otherwise report application error to {Constants.ApplicationDev}", 2);
            }
            res = true;
            return res;
        }
    }
}