﻿namespace Models
{
    public class MonthlyReportData
    {
        public List<EmployeeData> EmployeesData { get; init; } = [];

        public HashSet<string> ProjectIds { get; init; } = [];
    }
}