// This file is used by Code Analysis to maintain SuppressMessage
// attributes that are applied to this project.
// Project-level suppressions either have no target or are given
// a specific target and scoped to a namespace, type, member, etc.

using System.Diagnostics.CodeAnalysis;

[assembly: SuppressMessage(category: "Major Code Smell",
    checkId: "S6580:Use a format provider when parsing date and time",
    Justification = "DateTime format is not constant across files",
    Scope = "member",
    Target = "~M:Services.ReadService.ReadMonthlyReports(System.Collections.Generic.List{System.String},System.Object[],System.Int32)~Models.MonthlyReportData")]
[assembly: SuppressMessage("Major Code Smell",
    "S6580:Use a format provider when parsing date and time",
    Justification = "DateTime format is not constant across files",
    Scope = "member",
    Target = "~M:Services.ReadService.ReadMusterOptions(System.Collections.Generic.List{System.String})~Models.MusterOptionsDatas")]
[assembly: SuppressMessage("Major Code Smell",
    "S6580:Use a format provider when parsing date and time",
    Justification = "DateTime format is not constant across files",
    Scope = "member",
    Target = "~M:Services.ReadService.ReadPunchMovementReports(System.Collections.Generic.List{System.String})~Models.PunchMovementData")]