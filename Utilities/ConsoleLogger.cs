namespace Utilities
{
    public static class ConsoleLogger
    {
        private enum LogType
        {
            Normal,
            Info,
            Data,
            Warning,
            Error
        }

        public static void ExitApplication()
        {
            Log("Press any key to exit.", 1);
            _ = Console.ReadKey();
            Environment.Exit(0);
        }

        public static void Log(string message, int line = 0) => Log(LogType.Normal, message, line);
        
        public static void LogSameLine(string message, int line = 0) => Log(LogType.Normal, message, 0, line, true);

        public static void LogData(string data, int line = 0) => Log(LogType.Data, data, line);

        public static void LogDataSameLine(string data, int line = 0) => Log(LogType.Data, data, 0, line, true);
        
        public static void LogInfo(string message, int line = 0) => Log(LogType.Info, message, line);

        public static void LogInfoSameLine(string message, int line = 0) => Log(LogType.Info, message, 0, line, true);
        
        public static void LogWarning(string message, int line = 0) => Log(LogType.Warning, message, line);

        public static void LogWarningAndExit(string message, int line = 1)
        {
            LogLine(line);
            LogWarning(message);
            ExitApplication();
        }

        public static void LogError(string message, int line = 0) => Log(LogType.Error, message, line);
        
        public static void LogErrorAndExit(string message, int line = 1)
        {
            LogLine(line);
            LogError(message);
            ExitApplication();
        }

        public static void LogLine(int lines = 1)
        {
            for (int i = 0; i < lines; i++)
            {
                Console.WriteLine();
            }
        }

        private static void Log(LogType logType, string message, int lineBefore = 0, int lineAfter = 0, bool onSameLine = false)
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