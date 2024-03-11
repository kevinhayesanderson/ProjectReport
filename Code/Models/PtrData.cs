namespace Models
{
    public class PtrData
    {
        public Dictionary<string, TimeSpan> ProjectEfforts { get; init; } = [];

        public HashSet<string> ProjectIds { get; init; } = [];
    }
}