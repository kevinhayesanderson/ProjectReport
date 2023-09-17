namespace Models
{
    public class PtrData
    {
        public Dictionary<string, TimeSpan> ProjectEfforts { get; init; } = new Dictionary<string, TimeSpan>();

        public HashSet<string> ProjectIds { get; init; } = new HashSet<string>();
    }
}