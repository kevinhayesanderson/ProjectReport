namespace Models
{
    public class PunchData
    {
        public DateTime Date { get; init; }
        public List<TimeOnly> Punches { get; init; } = [];
        public TimeOnly FirstInTime { get; init; }
        public TimeOnly LastOutTime { get; init; }
        public TimeSpan TotalHours { get; init; }
        public TimeSpan WorkHours { get; init; }
        public TimeSpan BreakHours { get; init; }
        public bool IsLastOutNextDay { get; init; }
    }
}