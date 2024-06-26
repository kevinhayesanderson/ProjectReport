﻿using Microsoft.Extensions.FileSystemGlobbing;
using Models;
using Services;

namespace Actions
{
    [ActionName("CalculatePunchMovement")]
    internal class CalculatePunchMovementAction(string inputFolder, string cutOff) : Action
    {
        private List<string> _punchMovementFiles = [];
        public string InputFolder => inputFolder;

        public override bool Run()
        {
            var exportFolder = @$"{InputFolder}\Reports_{Time}";
            Logger.LogInfo("PunchMovement Files found:", 1);
            _punchMovementFiles.ForEach(pm => Logger.Log(new FileInfo(pm).Name));
            var punchMovementData = ReadService.ReadPunchMovementReports(_punchMovementFiles);
            DataService.CalculatePunchMovement(punchMovementData, cutOff);
            return ExportService.ExportPunchMovementSummaryReport(in exportFolder, Time, in punchMovementData);
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
                Matcher punchMovementMatcher = new();
                _ = punchMovementMatcher.AddInclude(Constants.PunchMovementPattern);
                _punchMovementFiles = punchMovementMatcher.GetResultsInFullPath(InputFolder).ToList();
                if (_punchMovementFiles.Count == 0)
                {
                    Logger.LogError($"No Punch Movement files found on {InputFolder}.");
                    res = false;
                }
            }
            return res;
        }
    }
}