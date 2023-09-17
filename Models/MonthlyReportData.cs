namespace Models
{
    public class MonthlyReportData
    {
        public List<EmployeeData> EmployeesData { get; init; } = new List<EmployeeData>();

        public HashSet<string> ProjectIds { get; init; } = new HashSet<string>();
    }
}