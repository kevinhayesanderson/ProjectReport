namespace Models
{
    public record Index(int Row = default, int Column = default);

    public static class Constants
    {
        public static class General
        {
            public const string ApplicationAdmin = "mohanraj.lvp@ambigai.net";
            public const int EmployeeIdLength = 5;
        }

        public static class MonthlyReport
        {
            public const string FileNamePattern = "*Monthly_Report*";
            public static readonly Index EmployeeNameIndex = new(Row: 3, Column: 2);
            public static readonly Index EmployeeIdIndex = new(Row: 4, Column: 2);
            public static readonly Index ActualAvailableHoursIndex = new(Row: 13);
            public static readonly Index LeavesRowIndex = new(Row: 14);
            public static readonly Index DataRowIndex = new(Row: 16);
            public const string SheetNamePattern = "MMM-yy";
            public static readonly Index InTimeIndex = new(Row: 10);
            public static readonly Index OutTimeIndex = new(Row: 12);
            public static readonly Index FirstDateIndex = new(Column: 5);
            public const string TimeNumberFormat = "[h]:mm";
        }

        public static class PTR
        {
            public const string FileNamePattern = "*ACS_PTR*";
        }

        public static class PunchMovement
        {
            public const string FileNamePattern = "*PunchMovement*";
            public static readonly Index HeadingsRowIndex = new(Row: 1);
            public const string ECodeColumnHeading = "ECode";
            public const string NameColumnHeading = "Name";
            public const string DateColumnHeading = "Date";
            public const string InColumnHeading = "IN";
            public const string OutColumnHeading = "OUT";
        }

        public static class MusterOptions
        {
            public const string FileNamePattern = "*Muster_Options*";
            public static readonly Index HeadingsRowIndex = new(Row: 3);
            public const string EmpCodeColumnHeading = "EmpCode";
            public const string NameColumnHeading = "Name";
            public const string DesignationColumnHeading = "Designation";
            public static readonly Index SerialNoIndex = new(Column: 0);
            public const int ShiftOffset = 1;
            public const int InTimeOffset = 2;
            public const int OutTimeOffset = 3;
            public const int MusterOffset = 4;
        }

        public static class AttendanceReport
        {
            public const string FileNamePattern = "*ACS_Attendance*";
            public const string SheetNamePattern = "MMM_yyyy";
            public static readonly Index EmpCodeIndex = new(Column: 1);
            public static readonly Index DateStartIndex = new(Column: 5);
            public const string TimeNumberFormat = "h:mm";
        }
    }
}