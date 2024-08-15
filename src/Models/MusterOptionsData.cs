namespace Models
{
    public class MusterOptionsDatas
    {
        public Dictionary<uint, MusterOptionsData> Datas { get; set; } = [];
    }

    public class MusterOptionsData
    {
        public string Name { get; init; } = string.Empty;
        public string Designation { get; init; } = string.Empty;
        public List<MusterOption> MusterOptions { get; init; } = [];
        public void AddMusterOptions(List<MusterOption> musterOptions) => MusterOptions.AddRange(musterOptions);
    }
}