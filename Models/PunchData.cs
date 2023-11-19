namespace Models
{
    public class PunchData
    {
        public DateTime Date { get; init; }
        public List<DateTime> Punches { get; init; } = [];
        public DateTime FirstInTime { get; init; }
        public DateTime LastOutTime { get; init; }
        public TimeSpan WorkHours { get; init; }
        public TimeSpan BreakHours { get; init; }
        public bool IsLastOutNextDay { get; init; }
    }
}