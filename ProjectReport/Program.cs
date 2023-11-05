using Models;
using ProjectReport.Actions;
using Services;
using System.Text;
using System.Text.RegularExpressions;
using Utilities;

namespace ProjectReport
{
    internal static partial class Program
    {
        private static readonly string _time = RemoveAllSymbols().Replace(DateTime.Now.ToString(), "_");

        private static void Main()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            Console.Title = $"Project Report PID:{Environment.ProcessId} {_time}";

            ConsoleLogger.LogInfo($"Running Project Report Application at {_time}");

            UserSettings userSettings = ReadService.GetUserSettings()!;

            var actions = IAction.InitiateActions(userSettings.Actions, _time);

            var res = IAction.ExecuteActions(actions);

            if (!res)
            {
                ConsoleLogger.LogErrorAndExit($"Application Error", 2);
            }

            ConsoleLogger.ExitApplication();
        }

        [GeneratedRegex("[^\\w\\d]")]
        private static partial Regex RemoveAllSymbols();
    }
}