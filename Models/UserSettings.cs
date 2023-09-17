namespace Models
{
    public class UserSettings
    {
        public string FinancialYear { get; set; } = string.Empty;

        public string Folder { get; set; } = string.Empty;

        public bool GenerateLeaveReport { get; set; }

        public List<string> MonthlyReportMonths { get; set; } = new List<string>();

        public int PtrBookingMonthCol { get; set; }

        public List<object> PtrBookingMonths { get; set; } = new List<object>();

        public List<double> PtrEffortCols { get; set; } = new List<double>();

        public int PtrProjectIdCol { get; set; }

        public string PtrSheetName { get; set; } = string.Empty;
    }
}