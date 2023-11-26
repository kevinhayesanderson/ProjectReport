namespace Actions
{
    [ActionName("InOutEntry")]
    internal class InOutEntryAction(string inputFolder) : Action
    {
        private List<string> _monthlyReports = [];
        public string InputFolder => inputFolder;
        public override bool Run()
        {
            bool res = true;
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
            }
            return res;
        }
    }
}