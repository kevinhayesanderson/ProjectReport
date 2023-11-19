using Microsoft.Extensions.FileSystemGlobbing;
using Services;
using Utilities;

namespace Actions
{
    [ActionName("CalculatePunchMovement")]
    internal class CalculatePunchMovementAction(bool run, string inputFolder, string time, ILogger logger, string cutOff, ReadService readService, DataService dataService, ExportService exportService) : IAction
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
            var punchMovementData = readService.ReadPunchMovementReports(PunchMovementFiles);
            dataService.CalculatePunchMovement(punchMovementData, cutOff);
            return exportService.ExportPunchMovementSummaryReport(in exportFolder, in punchMovementData);
        }
    }
}