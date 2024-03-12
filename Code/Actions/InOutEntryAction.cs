using Microsoft.Extensions.FileSystemGlobbing;
using Models;

namespace Actions
{
    [ActionName("InOutEntry")]
    internal class InOutEntryAction(string inputFolder) : Action
    {
        private List<string> _monthlyReports = [];
        private List<string> _musterOptionsReports = [];
        public string InputFolder => inputFolder;

        public override bool Run()
        {
            bool res = true;

            Logger.LogInfo("Muster Options files found found:", 1);
            _musterOptionsReports.ForEach(mor => Logger.Log(new FileInfo(mor).Name));
            Logger.LogInfo("Monthly reports found:", 1);
            _monthlyReports.ForEach(mr => Logger.Log(new FileInfo(mr).Name));

            var monthlyReportsData = _monthlyReports.Select(x => (DataService.ExtractEmployeeIdFromFileName(x), x)).ToList();

            MusterOptionsDatas _musterOptionsDatas = ReadService.ReadMusterOptions(_musterOptionsReports);

            if (_musterOptionsDatas != null && _musterOptionsDatas.Datas.Count > 0)
            {
                res = WriteService.WriteInOutEntry(monthlyReportsData, _musterOptionsDatas);
            }
            else
            {
                Logger.LogWarning($"Muster options data is empty, check if data is present, otherwise report application error to {Constants.ApplicationAdmin}", 2);
                return false;
            }
            return res;
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
                    Logger.LogError($"No Monthly reports found on {InputFolder}.");
                    res = false;
                }
                Matcher musterOptionMatcher = new();
                _ = musterOptionMatcher.AddInclude(Constants.MusterOptionsPattern);
                _musterOptionsReports = musterOptionMatcher.GetResultsInFullPath(InputFolder).ToList();
                if (_musterOptionsReports.Count == 0)
                {
                    Logger.LogError($"No Muster Options files found on {InputFolder}.");
                    res = false;
                }
            }
            return res;
        }
    }
}