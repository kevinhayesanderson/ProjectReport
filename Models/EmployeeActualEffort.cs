namespace Models
{
    public class EmployeeActualEffort
    {
        public TimeSpan ActualEffort { get; init; }

        public int Id { get; init; }

        public string Name { get; init; } = string.Empty;

        public string ProjectId { get; init; } = string.Empty;
    }
}
