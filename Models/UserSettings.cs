namespace Models
{
    using System.Text.Json.Serialization;

    public class UserSettings
    {
        [JsonPropertyName("Actions")]
        public required Action[] Actions { get; set; }
    }

    public class Action
    {
        [JsonPropertyName("Name")]
        public required string Name { get; set; }

        [JsonPropertyName("Run")]
        public required bool Run { get; set; }

        [JsonPropertyName("InputFolder")]
        public required string InputFolder { get; set; }

        [JsonPropertyName("MonthlyReportIdCol")]
        public int? MonthlyReportIdCol { get; set; }

        [JsonPropertyName("MonthlyReportMonths")]
        public object[]? MonthlyReportMonths { get; set; }

        [JsonPropertyName("PtrBookingMonthCol")]
        public int? PtrBookingMonthCol { get; set; }

        [JsonPropertyName("PtrBookingMonths")]
        public object[]? PtrBookingMonths { get; set; }

        [JsonPropertyName("PtrEffortCols")]
        public object[]? PtrEffortCols { get; set; }

        [JsonPropertyName("PtrProjectIdCol")]
        public int? PtrProjectIdCol { get; set; }

        [JsonPropertyName("PtrSheetName")]
        public string? PtrSheetName { get; set; }

        [JsonPropertyName("FinancialYear")]
        public string? FinancialYear { get; set; }

        [JsonPropertyName("Cut-off")]
        public string? CutOff { get; set; }
    }
}