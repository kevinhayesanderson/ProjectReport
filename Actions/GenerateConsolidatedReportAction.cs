using Microsoft.Extensions.FileSystemGlobbing;
using Models;
using Services;

namespace Actions
{
    [ActionName("GenerateConsolidatedReport")]
    internal class GenerateConsolidatedReportAction(string inputFolder,
                                                    object[] monthlyReportMonths,
                                                    int monthlyReportIdCol,
                                                    int ptrBookingMonthCol,
                                                    object[] ptrBookingMonths,
                                                    object[] ptrEffortCols,
                                                    int ptrProjectIdCol,
                                                    string ptrSheetName) : Action
    {
        private List<string> _monthlyReports = [];
        private List<string> _ptrFiles = [];

        public override bool Validate()
        {
            bool res = true;
            if (!Directory.Exists(inputFolder))
            {
                Logger.LogError($"Directory doesn't exist: {inputFolder}", 2);
                res = false;
            }
            else
            {
                _monthlyReports = Helper.GetMonthlyReports(inputFolder);
                if (_monthlyReports.Count == 0)
                {
                    Logger.LogError($"No Monthly reports found on {inputFolder}, needed monthly reports to generate consolidated report.");
                    res = false;
                }
                Matcher ptrMatcher = new();
                _ = ptrMatcher.AddInclude(Constants.PTRPattern);
                _ptrFiles = ptrMatcher.GetResultsInFullPath(inputFolder).ToList();
                if (_ptrFiles.Count == 0)
                {
                    Logger.LogWarning($"No PTR found on {inputFolder}, needed PTR to generate consolidated report.");
                    res = false;
                }
            }
            return res;
        }

        public override bool Run()
        {
            bool res = true;
            var _exportFolder = @$"{inputFolder}\Reports_{Time}";

            MonthlyReportData? _monthlyReportData = default;

            Logger.LogInfo("Monthly reports found:", 1);
            _monthlyReports.ForEach(mr => Logger.Log(new FileInfo(mr).Name));
            _monthlyReportData = ReadService.ReadMonthlyReports(_monthlyReports, monthlyReportMonths, monthlyReportIdCol);
            ExportService.ExportMonthlyReportInter(ref _monthlyReportData, Time, in _exportFolder);

            PtrData? _ptrData = default;

            Logger.LogInfo("PTR's found:", 2);
            _ptrFiles.ForEach(ptr => Logger.Log(new FileInfo(ptr).Name));
            _ptrData = ReadService.ReadPtr(_ptrFiles, ptrBookingMonthCol, ptrBookingMonths, ptrEffortCols, ptrProjectIdCol, ptrSheetName);
            ExportService.ExportPtrInter(ref _ptrData, Time, in _exportFolder);

            if (_monthlyReportData != null && _ptrData != null)
            {
                List<ConsolidatedData> ConsolidatedData = DataService.Consolidate(_ptrData, _monthlyReportData);
                if (ConsolidatedData.Count > 0)
                {
                    ExportService.ExportConsolidateData(in ConsolidatedData, ref _monthlyReportData, Time, in _exportFolder);
                }
                else
                {
                    Logger.LogWarning($"Consolidated data is empty, modify filter conditions or check if data is present, otherwise report application error to {Constants.ApplicationAdmin}", 2);
                    return false;
                }
            }
            else
            {
                Logger.LogWarning($"Either PTR or Monthly report data is empty, modify filter conditions or check if data is present, otherwise report application error to {Constants.ApplicationAdmin}", 2);
                return false;
            }
            return res;
        }
    }
}