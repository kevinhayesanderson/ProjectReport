using Services;

namespace ProjectReport.Actions
{
    [SettingName("GenerateLeaveReport")]
    internal class GenerateLeaveReportAction : IAction
    {
        private readonly string _fy;
        private readonly string _time;

        public GenerateLeaveReportAction(bool run, string inputFolder, string time, string fy)
        {
            (Run, InputFolder) = (run, inputFolder);
            (_time, _fy) = (time, fy);
        }

        public string InputFolder { get; }
        public bool Run { get; }

        public bool Execute()
        {
            var _exportFolder = @$"{InputFolder}\Reports_{_time}";
            List<string> monthlyReports = Helper.GetMonthlyReports(InputFolder);
            ExportService.ExportLeaveReport(in monthlyReports, _fy, in _exportFolder);
            bool res = true;
            return res;
        }
    }
}