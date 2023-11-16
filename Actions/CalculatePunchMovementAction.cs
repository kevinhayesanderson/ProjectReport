using Microsoft.Extensions.FileSystemGlobbing;
using Utilities;

namespace Actions
{
    [ActionName("CalculatePunchMovement")]
    internal class CalculatePunchMovementAction(bool run, string inputFolder, string time, ILogger logger) : IAction
    {
        public string InputFolder => inputFolder;

        public bool Run => run;

        public bool Execute()
        {
            var exportFolder = @$"{InputFolder}\Reports_{time}";
            Matcher punchMovementMatcher = new();
            _ = punchMovementMatcher.AddInclude(Constants.PunchMovementPattern);
            List<string> PunchMovementFiles = punchMovementMatcher.GetResultsInFullPath(InputFolder).ToList();
            logger.LogInfo("PunchMovement Files found:", 1);
            PunchMovementFiles.ForEach(pm => logger.Log(new FileInfo(pm).Name));
            bool res = true;
            return res;
        }
    }
}