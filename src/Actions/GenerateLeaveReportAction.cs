using Models;
using Services;
using Utilities;

namespace Actions
{
    [ActionName("GenerateLeaveReport")]
    internal class GenerateLeaveReportAction(string fy) : Action
    {
        private List<string> _monthlyReports = [];

        public override void Init()
        {
            _monthlyReports = Helper.GetReports(InputFolder, Constants.MonthlyReport.FileNamePattern).ToList();
        }

        public override bool Run()
        {
            Logger.LogFileNames(_monthlyReports, "Monthly reports found:");

            return ExportService.ExportLeaveReport(in _monthlyReports, fy, ExportFolder);
        }

        public override bool Validate()
        {
            bool res = ValidateDirectory(InputFolder);

            res = res && ValidateReports(_monthlyReports, $"No Monthly Report files with naming pattern {Constants.MonthlyReport.FileNamePattern} found on {InputFolder}.");

            return res;
        }
    }
}