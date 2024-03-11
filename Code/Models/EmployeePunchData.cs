namespace Models
{
    public class EmployeePunchData
    {
        public int Id { get; init; }
        public string Name { get; init; } = string.Empty;
        public List<PunchData> PunchDatas { get; init; } = [];
        public TimeSpan TotalAvailableHours { get; set; }
        public TimeSpan TotalWorkHours { get; set; }
        public TimeSpan TotalBreakHours { get; set; }
    }
}