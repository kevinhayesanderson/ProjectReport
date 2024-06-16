namespace Models
{
    public class PunchData
    {
        public TimeSpan AvailableHours { get; init; }
        public TimeSpan BreakHours { get; init; }
        public DateTime Date { get; init; }
        public TimeOnly FirstInTime { get; init; }
        public bool IsLastOutNextDay { get; init; }
        public TimeOnly LastOutTime { get; init; }
        public List<TimeOnly> Punches { get; init; } = [];
        public TimeSpan WorkHours { get; init; }
    }
}