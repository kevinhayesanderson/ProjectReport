using Microsoft.Extensions.FileSystemGlobbing;

namespace Utilities
{
    public static class Helper
    {
        public static IEnumerable<string> GetReports(string inputFolder, string pattern) => new Matcher().AddInclude(pattern).GetResultsInFullPath(inputFolder);
    }
}