namespace Models
{
    public class EmployeePunchData
    {
        public int Id { get; init; }
        public string Name { get; init; } = string.Empty;
        public List<PunchData> PunchDatas { get; init; } = [];
    }
}