namespace Models
{
    public class EmployeeData
    {
        public TimeSpan ActualAvailableHours { get; init; }

        public int Id { get; init; }

        public string Name { get; init; } = string.Empty;

        public Dictionary<string, TimeSpan> ProjectData { get; init; } = [];

        public int TotalLeaves { get; init; }

        public TimeSpan TotalProjectHours { get; set; }
    }
}