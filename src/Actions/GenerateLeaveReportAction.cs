using Models;
using Services;
using Utilities;

namespace Actions
{
    [ActionName("GenerateLeaveReport")]
    internal class GenerateLeaveReportAction(string inputFolder, string fy) : Action
    {
        private List<string> _monthlyReports = [];

        private string _exportFolder = string.Empty;

        public override void Init()
        {
            _monthlyReports = Helper.GetReports(inputFolder, Constants.MonthlyReport.FileNamePattern).ToList();

            _exportFolder = @$"{inputFolder}\Reports_{Time}";
        }

        public override bool Run()
        {
            Logger.LogFileNames(_monthlyReports, "Monthly reports found:");

            return ExportService.ExportLeaveReport(in _monthlyReports, fy, in _exportFolder);
        }

        public override bool Validate()
        {
            bool res = ValidateDirectory(inputFolder);

            res = res && ValidateReports(_monthlyReports, $"No Monthly Report files with naming pattern {Constants.MonthlyReport.FileNamePattern} found on {inputFolder}.");

            return res;
        }
    }
}