using Microsoft.Extensions.FileSystemGlobbing;
using Models;
using Services;
using Utilities;

namespace Actions
{
    [ActionName("GenerateConsolidatedReport")]
    internal class GenerateConsolidatedReportAction(bool run,
                                                    string inputFolder,
                                                    string time,
                                                    ILogger logger,
                                                    object[] monthlyReportMonths,
                                                    int monthlyReportIdCol,
                                                    int ptrBookingMonthCol,
                                                    object[] ptrBookingMonths,
                                                    object[] ptrEffortCols,
                                                    int ptrProjectIdCol,
                                                    string ptrSheetName, DataService dataService, ReadService readService, ExportService exportService) : IAction
    {
        public string InputFolder => inputFolder;
        public bool Run => run;

        public bool Execute()
        {
            bool res = true;
            var _exportFolder = @$"{InputFolder}\Reports_{time}";

            MonthlyReportData? _monthlyReportData = default;
            List<string> monthlyReports = Helper.GetMonthlyReports(InputFolder);
            if (monthlyReports.Count > 0)
            {
                logger.LogInfo("Monthly reports found:", 1);
                monthlyReports.ForEach(mr => logger.Log(new FileInfo(mr).Name));
                _monthlyReportData = readService.ReadMonthlyReports(monthlyReports, monthlyReportMonths, monthlyReportIdCol);
                exportService.ExportMonthlyReportInter(ref _monthlyReportData, in time, in _exportFolder);
            }
            else
            {
                logger.LogWarning($"No Monthly reports found on {InputFolder}, needed monthly reports to generate consolidated report or leave report.");
                return false;
            }

            PtrData? _ptrData = default;
            Matcher ptrMatcher = new();
            _ = ptrMatcher.AddInclude(Constants.PTRPattern);
            List<string> ptrFiles = ptrMatcher.GetResultsInFullPath(InputFolder).ToList();
            if (ptrFiles.Count > 0)
            {
                logger.LogInfo("PTR's found:", 2);
                ptrFiles.ForEach(ptr => logger.Log(new FileInfo(ptr).Name));
                _ptrData = readService.ReadPtr(ptrFiles, ptrBookingMonthCol, ptrBookingMonths, ptrEffortCols, ptrProjectIdCol, ptrSheetName);
                exportService.ExportPtrInter(ref _ptrData, in time, in _exportFolder);
            }
            else
            {
                logger.LogWarning($"No PTR found on {InputFolder}, needed PTR to generate consolidated report.");
                return false;
            }

            if (_monthlyReportData != null && _ptrData != null)
            {
                List<ConsolidatedData> ConsolidatedData = dataService.Consolidate(_ptrData, _monthlyReportData);
                if (ConsolidatedData.Count > 0)
                {
                    exportService.ExportConsolidateData(in ConsolidatedData, ref _monthlyReportData, in time, in _exportFolder);
                }
                else
                {
                    logger.LogWarning($"Consolidated data is empty, modify filter conditions or check if data is present, otherwise report application error to {Constants.ApplicationAdmin}", 2);
                    return false;
                }
            }
            else
            {
                logger.LogWarning($"Either PTR or Monthly report data is empty, modify filter conditions or check if data is present, otherwise report application error to {Constants.ApplicationAdmin}", 2);
                return false;
            }
            return res;
        }
    }
}