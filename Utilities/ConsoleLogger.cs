namespace Utilities
{
    public class ConsoleLogger : ILogger
    {
        private enum LogType
        {
            Normal,
            Info,
            Data,
            Warning,
            Error
        }

        public void ExitApplication()
        {
            Log("Press any key to exit.", 1);
            _ = Console.ReadKey();
            Environment.Exit(0);
        }

        public void ExitApplication(string exitMessage)
        {
            Log(exitMessage, 1);
            ExitApplication();
        }

        public void Log(string message, int line = 0) => Log(LogType.Normal, message, line);

        public void LogChar(char character, int length)
        {
            for (int i = 0; i < length; i++)
            {
                Console.Write(character);
            }
        }

        public void LogData(string data, int line = 0) => Log(LogType.Data, data, line);

        public void LogDataSameLine(string data, int line = 0) => Log(LogType.Data, data, 0, line, true);

        public void LogError(string message, int line = 0) => Log(LogType.Error, message, line);

        public void LogErrorAndExit(string message, int line = 1)
        {
            LogLine(line);
            LogError(message);
            ExitApplication();
        }

        public void LogInfo(string message, int line = 0) => Log(LogType.Info, message, line);

        public void LogInfoSameLine(string message, int line = 0) => Log(LogType.Info, message, 0, line, true);

        public void LogLine(int lines = 1)
        {
            for (int i = 0; i < lines; i++)
            {
                Console.WriteLine();
            }
        }

        public void LogSameLine(string message, int line = 0) => Log(LogType.Normal, message, 0, line, true);

        public void LogWarning(string message, int line = 0) => Log(LogType.Warning, message, line);

        public void LogWarningAndExit(string message, int line = 1)
        {
            LogLine(line);
            LogWarning(message);
            ExitApplication();
        }

        private void Log(LogType logType, string message, int lineBefore = 0, int lineAfter = 0, bool onSameLine = false)
        {
            Console.BackgroundColor = ConsoleColor.Black;
            Console.ForegroundColor = logType switch
            {
                LogType.Normal => ConsoleColor.White,
                LogType.Info => ConsoleColor.Cyan,
                LogType.Data => ConsoleColor.Green,
                LogType.Warning => ConsoleColor.Yellow,
                LogType.Error => ConsoleColor.Red,
                _ => ConsoleColor.White,
            };
            LogLine(lineBefore);
            if (onSameLine)
            {
                Console.Write(message);
            }
            else
            {
                Console.WriteLine(message);
            }
            LogLine(lineAfter);
            Console.ResetColor();
        }
    }
}