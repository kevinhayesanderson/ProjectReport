using Microsoft.Extensions.FileSystemGlobbing;
using Utilities;

namespace Actions
{
    [SettingName("CalculatePunchMovement")]
    internal class CalculatePunchMovementAction : IAction
    {
        private readonly string _time;

        public CalculatePunchMovementAction(bool run, string inputFolder, string time)
        {
            Run = run;
            InputFolder = inputFolder;
            _time = time;
        }

        public string InputFolder { get; }
        public bool Run { get; }

        public bool Execute()
        {
            Matcher punchMovementMatcher = new();
            _ = punchMovementMatcher.AddInclude(Constants.PunchMovementPattern);
            List<string> PunchMovementFiles = punchMovementMatcher.GetResultsInFullPath(InputFolder).ToList();
            ConsoleLogger.LogInfo("PunchMovement Files found:", 1);
            PunchMovementFiles.ForEach(pm => ConsoleLogger.Log(new FileInfo(pm).Name));
            bool res = true;
            return res;
        }
    }
}