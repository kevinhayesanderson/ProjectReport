namespace Models
{
    public class ConsolidatedData
    {
        public List<EmployeeActualEffort> EmployeeActualEffort { get; init; } = [];

        public string ProjectId { get; init; } = string.Empty;

        public TimeSpan TotalEffort { get; init; }
    }
}