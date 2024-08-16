using Models;
using Utilities;

namespace Actions
{
    [ActionName("MonthlyReportInOutEntry")]
    internal class MonthlyReportInOutEntryAction() : Action
    {
        private List<string> _monthlyReports = [];
        private List<string> _musterOptionsReports = [];

        public override void Init()
        {
            _monthlyReports = Helper.GetReports(InputFolder, Constants.MonthlyReport.FileNamePattern).ToList();

            _musterOptionsReports = Helper.GetReports(InputFolder, Constants.MusterOptions.FileNamePattern).ToList();
        }

        public override bool Run()
        {
            Logger.LogFileNames(_musterOptionsReports, "Muster Options files found found:");

            Logger.LogFileNames(_monthlyReports, "Monthly reports found:");

            var monthlyReportsData = _monthlyReports.Select(x => (Services.DataService.ExtractEmployeeIdFromFileName(x), x)).ToList();

            var musterOptionsDatas = ReadService.ReadMusterOptions(_musterOptionsReports);

            if (musterOptionsDatas != null && musterOptionsDatas.Datas.Count > 0)
            {
                return WriteService.WriteMonthlyReportInOutEntry(monthlyReportsData, musterOptionsDatas);
            }
            else
            {
                Logger.LogWarning($"Muster options data is empty, check if data is present, otherwise report application error to {Constants.General.ApplicationAdmin}", 2);
                return false;
            }
        }

        public override bool Validate()
        {
            bool res = ValidateDirectory(InputFolder);

            res = res && ValidateReports(_monthlyReports, $"No Monthly Report files with naming pattern {Constants.MonthlyReport.FileNamePattern} found on {InputFolder}.");

            res = res && ValidateReports(_musterOptionsReports, $"No Muster Options files with naming pattern {Constants.MusterOptions.FileNamePattern} found on {InputFolder}.");

            return res;
        }
    }
}