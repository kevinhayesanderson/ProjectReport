using Models;
using Services;
using Utilities;

namespace Actions
{
    [ActionName("CalculatePunchMovement")]
    internal class CalculatePunchMovementAction(string cutOff) : Action
    {
        private List<string> _punchMovementFiles = [];

        public override void Init()
        {
            _punchMovementFiles = Helper.GetReports(InputFolder, Constants.PunchMovement.FileNamePattern).ToList();
        }

        public override bool Run()
        {
            Logger.LogFileNames(_punchMovementFiles, "PunchMovement files found:");

            var punchMovementData = ReadService.ReadPunchMovementReports(_punchMovementFiles);

            DataService.CalculatePunchMovement(punchMovementData, cutOff);

            return ExportService.ExportPunchMovementSummaryReport(ExportFolder, Time, in punchMovementData);
        }

        public override bool Validate()
        {
            bool res = ValidateDirectory(InputFolder);

            res = res && ValidateReports(_punchMovementFiles, $"No Punch Movement files with naming pattern {Constants.PunchMovement.FileNamePattern} found on {InputFolder}");

            return res;
        }
    }
}