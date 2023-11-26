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
            
            var _exportFolder = @$"{inputFolder}\Reports_{Time}";

            MusterOptionsData? _musterOptionsData = default;
            Logger.LogInfo("Muster Options files found found:", 1);
            _musterOptionsReports.ForEach(mr => Logger.Log(new FileInfo(mr).Name));
            _musterOptionsData = ReadService.ReadMusterOptions(_musterOptionsReports);

            if(_musterOptionsData != null)
            {
                res = WriteService.WriteInOutEntry(_monthlyReports, _musterOptionsData, _exportFolder);
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