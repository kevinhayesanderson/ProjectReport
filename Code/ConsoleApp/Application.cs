using Services;
using System.Text;
using System.Text.RegularExpressions;
using Utilities;

namespace ConsoleApp
{
    internal class Application(ILogger logger, DataService dataService, ReadService readService, WriteService writeService, ExportService exportService)
    {
        public void Run()
        {
            string time = new Regex("[^\\w\\d]").Replace(DateTime.Now.ToString(), "_");

            CancellationTokenSource cts = new();

            Console.Clear();

            Console.CancelKeyPress += (sender, args) => ExitHandler(cts);

            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            Console.Title = $"Project Report Application PID:{Environment.ProcessId}";

            logger.LogInfo($"Running Project Report Application at Date_Time:{time}");

            var userSettings = readService.GetUserSettings();

            if (userSettings != null)
            {
                Actions.Action.Init(time, logger, dataService, readService, writeService, exportService);

                var res = Actions.Action.ExecuteActions(userSettings!.Actions, cts.Token);

                if (!res)
                {
                    ExitApplication($"Application Error", 2);
                }
            }

            ExitApplication("Exiting application.");
        }

        private void ExitHandler(CancellationTokenSource cts)
        {
            cts.Cancel();
            logger.LogWarning("Application canceled by user");
            ExitApplication("Exiting application.");
        }

        public void ExitApplication(string exitMessage = "", int line = 1)
        {
            if (!string.IsNullOrEmpty(exitMessage))
                logger.Log(exitMessage, line);
            Console.WriteLine("Press any key to exit...");
            _ = Console.ReadKey(false);
            Environment.Exit(0);
        }
    }
}