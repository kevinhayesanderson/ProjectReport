namespace Actions
{
    [ActionName("InOutEntry")]
    internal class InOutEntryAction(string inputFolder) : Action
    {
        public string InputFolder => inputFolder;
        private List<string> _monthlyReports = [];
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

        public override bool Run()
        {
            bool res = true;
            return res;
        }
    }
}