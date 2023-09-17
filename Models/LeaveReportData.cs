namespace Models
{
    public class LeaveReportData
    {
        public string EmployeeId { get; init; } = string.Empty;

        public Dictionary<string, int?> Leaves { get; init; } = new Dictionary<string, int?>();

        public string Name { get; init; } = string.Empty;

        public int? TotalLeaves { get; init; }
    }
}