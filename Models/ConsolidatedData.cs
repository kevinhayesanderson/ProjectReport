﻿namespace Models
{
    public class ConsolidatedData
    {
        public List<EmployeeActualEffort> EmployeeActualEffort { get; init; } = new List<EmployeeActualEffort>();

        public string ProjectId { get; init; } = string.Empty;

        public TimeSpan TotalEffort { get; init; }
    }
}