using Microsoft.Extensions.FileSystemGlobbing;

namespace ProjectReport.Actions
{
    internal static class Helper
    {
        public static List<string> GetMonthlyReports(string inputFolder)
        {
            Matcher monthlyReportMatcher = new();
            _ = monthlyReportMatcher.AddInclude(Constants.MonthlyReportPattern);
            return monthlyReportMatcher.GetResultsInFullPath(inputFolder).ToList();
        }
    }
}