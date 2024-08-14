namespace Models
{
    public class MusterOption
    {
        public DateTime Date { get; init; }
        public TimeOnly? InTime { get; init; }
        public TimeOnly? OutTime { get; init; }
        public string Shift { get; init; } = string.Empty;
        public string Muster { get; init; } = string.Empty;
    }
}