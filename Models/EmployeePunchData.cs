namespace Models
{
    public class EmployeePunchData
    {
        public int Id { get; init; }
        public string Name { get; init; } = string.Empty;
        public DateTime Date { get; init; }
        public List<DateTime> Punches { get; init; } = [];
        public DateTime FirstInTime { get; init; }
        public DateTime LastInTime { get; init; }
        public DateTime FirstOutTime { get; init; }
        public DateTime LastOutTime { get; init; }
        public TimeSpan TotalWorkInHours { get; init; }
        public TimeSpan TotalBreakInHours { get; init; }
    }
}