namespace Models
{
    public class MusterOptionsData
    {
        public int Id { get; init; }
        public string Name { get; init; } = string.Empty;
        public List<MusterOption> MusterOptions { get; init; } = [];
    }
}