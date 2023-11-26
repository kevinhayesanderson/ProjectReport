using Services;

namespace Actions
{
    [ActionName("GenerateLeaveReport")]
    internal class GenerateLeaveReportAction(string inputFolder, string fy) : Action
    {
        private List<string> _monthlyReports = [];
        public string InputFolder => inputFolder;

        public override bool Run()
        {
            var _exportFolder = @$"{InputFolder}\Reports_{Time}";
            Logger.LogInfo("Monthly reports found:", 1);
            _monthlyReports.ForEach(mr => Logger.Log(new FileInfo(mr).Name));
            return ExportService.ExportLeaveReport(in _monthlyReports, fy, in _exportFolder);
        }

        public override bool Validate()
        {
            bool res = true;
            if (!Directory.Exists(InputFolder))
            {
                Logger.LogError($"Directory doesn't exist: {InputFolder}", 2);
                res = false;
            }
            else
            {
                _monthlyReports = Helper.GetMonthlyReports(InputFolder);
                if (_monthlyReports.Count == 0)
                {
                    Logger.LogError($"No Monthly reports found on {InputFolder}, needed monthly reports to generate leave report.");
                    res = false;
                }
            }
            return res;
        }
    }
}