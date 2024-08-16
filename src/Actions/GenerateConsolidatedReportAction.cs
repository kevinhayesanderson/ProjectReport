using Models;
using Services;
using Utilities;

namespace Actions
{
    [ActionName("GenerateConsolidatedReport")]
    internal class GenerateConsolidatedReportAction(
        object[] monthlyReportMonths, int monthlyReportIdCol, int ptrBookingMonthCol, object[] ptrBookingMonths, object[] ptrEffortCols, int ptrProjectIdCol, string ptrSheetName) : Action
    {
        private List<string> _monthlyReports = [];
        private List<string> _ptrFiles = [];

        public override void Init()
        {
            _monthlyReports = Helper.GetReports(InputFolder, Constants.MonthlyReport.FileNamePattern).ToList();

            _ptrFiles = Helper.GetReports(InputFolder, Constants.PTR.FileNamePattern).ToList();
        }

        public override bool Run()
        {
            Logger.LogFileNames(_monthlyReports, "Monthly reports found:");

            var _monthlyReportData = ReadService.ReadMonthlyReports(_monthlyReports, monthlyReportMonths, monthlyReportIdCol);
            ExportService.ExportMonthlyReportInter(ref _monthlyReportData, Time, ExportFolder);

            Logger.LogFileNames(_ptrFiles, "PTR's found:");

            var _ptrData = ReadService.ReadPtr(_ptrFiles, ptrBookingMonthCol, ptrBookingMonths, ptrEffortCols, ptrProjectIdCol, ptrSheetName);
            ExportService.ExportPtrInter(ref _ptrData, Time, ExportFolder);

            if (_monthlyReportData != null && _ptrData != null)
            {
                var ConsolidatedData = DataService.Consolidate(_ptrData, _monthlyReportData);

                if (ConsolidatedData.Count > 0)
                {
                    return ExportService.ExportConsolidateData(in ConsolidatedData, ref _monthlyReportData, Time, ExportFolder);
                }
                else
                {
                    Logger.LogWarning($"Consolidated data is empty, modify filter conditions or check if data is present, otherwise report application error to {Constants.General.ApplicationAdmin}", 2);
                    return false;
                }
            }
            else
            {
                Logger.LogWarning($"Either PTR or Monthly report data is empty, modify filter conditions or check if data is present, otherwise report application error to {Constants.General.ApplicationAdmin}", 2);
                return false;
            }
        }

        public override bool Validate()
        {
            bool res = ValidateDirectory(InputFolder);

            res = res && ValidateReports(_monthlyReports, $"No Monthly Report files with naming pattern {Constants.MonthlyReport.FileNamePattern} found on {InputFolder}.");

            res = res && ValidateReports(_ptrFiles, $"No PTR files with naming pattern {Constants.PTR.FileNamePattern} found on {InputFolder}.");

            return res;
        }
    }
}