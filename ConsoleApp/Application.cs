using Actions;
using Services;
using System.Text;
using System.Text.RegularExpressions;
using Utilities;

namespace ConsoleApplication
{
    internal class Application(ILogger logger, DataService dataService, ReadService readService, WriteService writeService, ExportService exportService)
    {
        public void Run()
        {
            string time = new Regex("[^\\w\\d]").Replace(DateTime.Now.ToString(), "_");

            CancellationTokenSource cts = new();

            Console.Clear();

            AppDomain.CurrentDomain.ProcessExit += (sender, args) => ExitHandler(cts);

            Console.CancelKeyPress += (sender, args) => ExitHandler(cts);

            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            Console.Title = $"Project Report Application PID:{Environment.ProcessId} Date_Time:{time}";

            logger.LogInfo($"Running Project Report Application at Date_Time:{time}");

            var userSettings = readService.GetUserSettings();

            if (userSettings != null)
            {
                var res = IAction.ExecuteActions(userSettings!.Actions, time, logger, dataService, readService, writeService, exportService, cts.Token);

                if (!res)
                {
                    logger.LogErrorAndExit($"Application Error", 2);
                }
            }

            logger.ExitApplication("Exiting application.");
        }

        private void ExitHandler(CancellationTokenSource cts)
        {
            cts.Cancel();
            logger.ExitApplication("Exiting...");
        }
    }
}