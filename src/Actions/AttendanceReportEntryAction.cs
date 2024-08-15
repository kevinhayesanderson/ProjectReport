﻿using Models;
using Utilities;

namespace Actions
{
    [ActionName("AttendanceReportEntry")]
    internal class AttendanceReportEntryAction(string inputFolder) : Action
    {
        private List<string> _attendanceReports = [];
        private List<string> _musterOptionsReports = [];

        public override void Init()
        {
            _attendanceReports = Helper.GetReports(inputFolder, Constants.AttendanceReportPattern).ToList();

            _musterOptionsReports = Helper.GetReports(inputFolder, Constants.MusterOptionsPattern).ToList();
        }

        public override bool Run()
        {
            Logger.LogFileNames(_musterOptionsReports, "Muster Options files found found:");

            Logger.LogFileNames(_attendanceReports, "Attendance Report files found:");

            var musterOptionsDatas = ReadService.ReadMusterOptions(_musterOptionsReports);

            if (musterOptionsDatas != null && musterOptionsDatas.Datas.Count > 0)
            {
                return WriteService.WriteAttendanceReportEntry(_attendanceReports, musterOptionsDatas);
            }
            else
            {
                Logger.LogWarning($"Muster options data is empty, check if data is present, otherwise report application error to {Constants.ApplicationAdmin}", 2);
                return false;
            }
        }

        public override bool Validate()
        {
            bool res = ValidateDirectory(inputFolder);

            res = res && ValidateReports(_attendanceReports, $"No Attendance Report files with naming pattern {Constants.AttendanceReportPattern} found on {inputFolder}.");

            res = res && ValidateReports(_musterOptionsReports, $"No Muster Options files with naming pattern {Constants.MusterOptionsPattern} found on {inputFolder}.");

            return res;
        }
    }
}