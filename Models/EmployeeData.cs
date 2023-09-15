namespace Models
{
    public class EmployeeData
    {
        public int Id { get; init; }

        public string Name { get; init; } = string.Empty;

        public Dictionary<string, TimeSpan> ProjectTime { get; init; } = new Dictionary<string, TimeSpan>();
    }
}
