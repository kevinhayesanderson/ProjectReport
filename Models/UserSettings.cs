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
        public bool Run { get; set; }

        [JsonPropertyName("InputFolder")]
        public required string InputFolder { get; set; }

        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        [JsonPropertyName("MonthlyReportIdCol")]
        public int? MonthlyReportIdCol { get; set; }

        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        [JsonPropertyName("MonthlyReportMonths")]
        public object[]? MonthlyReportMonths { get; set; }

        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        [JsonPropertyName("PtrBookingMonthCol")]
        public int? PtrBookingMonthCol { get; set; }

        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        [JsonPropertyName("PtrBookingMonths")]
        public object[]? PtrBookingMonths { get; set; }

        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        [JsonPropertyName("PtrEffortCols")]
        public object[]? PtrEffortCols { get; set; }

        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        [JsonPropertyName("PtrProjectIdCol")]
        public int? PtrProjectIdCol { get; set; }

        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        [JsonPropertyName("PtrSheetName")]
        public string? PtrSheetName { get; set; }

        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        [JsonPropertyName("FinancialYear")]
        public string? FinancialYear { get; set; }
    }
}