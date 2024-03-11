using System.Text.Json.Serialization;

namespace Models
{
    public class UserSettings
    {
        [JsonPropertyName("Actions")]
        public Action[] Actions { get; set; }
    }

    public class Action
    {
        [JsonPropertyName("Name")]
        public string Name { get; set; }

        [JsonPropertyName("Run")]
        public bool Run { get; set; }

        [JsonPropertyName("InputFolder")]
        public string InputFolder { get; set; }

        [JsonPropertyName("MonthlyReportIdCol")]
        public int MonthlyReportIdCol { get; set; } = -1;

        [JsonPropertyName("MonthlyReportMonths")]
        public object[] MonthlyReportMonths { get; set; } = [];

        [JsonPropertyName("PtrBookingMonthCol")]
        public int PtrBookingMonthCol { get; set; } = -1;

        [JsonPropertyName("PtrBookingMonths")]
        public object[] PtrBookingMonths { get; set; } = [];

        [JsonPropertyName("PtrEffortCols")]
        public object[] PtrEffortCols { get; set; } = [];

        [JsonPropertyName("PtrProjectIdCol")]
        public int PtrProjectIdCol { get; set; } = -1;

        [JsonPropertyName("PtrSheetName")]
        public string PtrSheetName { get; set; } = string.Empty;

        [JsonPropertyName("FinancialYear")]
        public string FinancialYear { get; set; } = string.Empty;

        [JsonPropertyName("CutOff")]
        public string CutOff { get; set; } = string.Empty;
    }
}