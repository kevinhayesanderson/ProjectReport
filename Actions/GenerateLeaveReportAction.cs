using Services;
using Utilities;

namespace Actions
{
    [ActionName("GenerateLeaveReport")]
    internal class GenerateLeaveReportAction(bool run, string inputFolder, string time, ILogger logger, string fy, ExportService exportService) : IAction
    {
        public string InputFolder => inputFolder;
        public bool Run => run;

        public bool Execute()
        {
            var _exportFolder = @$"{InputFolder}\Reports_{time}";
            List<string> monthlyReports = Helper.GetMonthlyReports(InputFolder);
            exportService.ExportLeaveReport(in monthlyReports, fy, in _exportFolder);
            bool res = true;
            return res;
        }
    }
}