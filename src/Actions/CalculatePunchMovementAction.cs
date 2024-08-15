using Models;
using Services;
using Utilities;

namespace Actions
{
    [ActionName("CalculatePunchMovement")]
    internal class CalculatePunchMovementAction(string inputFolder, string cutOff) : Action
    {
        private List<string> _punchMovementFiles = [];
        private string _exportFolder = string.Empty;

        public override void Init()
        {
            _punchMovementFiles = Helper.GetReports(inputFolder, Constants.PunchMovementPattern).ToList();
            _exportFolder = @$"{inputFolder}\Reports_{Time}";
        }

        public override bool Run()
        {
            Logger.LogFileNames(_punchMovementFiles, "PunchMovement files found:");

            var punchMovementData = ReadService.ReadPunchMovementReports(_punchMovementFiles);

            DataService.CalculatePunchMovement(punchMovementData, cutOff);

            return ExportService.ExportPunchMovementSummaryReport(in _exportFolder, Time, in punchMovementData);
        }

        public override bool Validate()
        {
            bool res = ValidateDirectory(inputFolder);

            res = res && ValidateReports(_punchMovementFiles, $"No Punch Movement files with naming pattern {Constants.PunchMovementPattern} found on {inputFolder}");

            return res;
        }
    }
}