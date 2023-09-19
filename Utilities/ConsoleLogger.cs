namespace Utilities
{
    public static class ConsoleLogger
    {
        public static void ExitApplication()
        {
            Log("Press any key to exit.", 1);
            _ = Console.ReadKey();
            Environment.Exit(0);
        }

        public static void Log(string message, int line = 0)
        {
            Console.BackgroundColor = ConsoleColor.Black;
            Console.ForegroundColor = ConsoleColor.White;
            LogLine(line);
            Console.WriteLine(message);
            Console.ResetColor();
        }

        public static void LogData(string data, int line = 0)
        {
            Console.BackgroundColor = ConsoleColor.Black;
            Console.ForegroundColor = ConsoleColor.Green;
            LogLine(line);
            Console.WriteLine(data);
            Console.ResetColor();
        }

        public static void LogDataSameLine(string data, int line = 0)
        {
            Console.BackgroundColor = ConsoleColor.Black;
            Console.ForegroundColor = ConsoleColor.Green;
            Console.Write(data);
            LogLine(line);
            Console.ResetColor();
        }

        public static void LogErrorAndExit(string message, int line = 1)
        {
            LogLine(line);
            LogError(message);
            ExitApplication();
        }

        public static void LogInfo(string message, int line = 0)
        {
            Console.BackgroundColor = ConsoleColor.Black;
            Console.ForegroundColor = ConsoleColor.Cyan;
            LogLine(line);
            Console.WriteLine(message);
            Console.ResetColor();
        }

        public static void LogInfoSameLine(string message, int line = 0)
        {
            Console.BackgroundColor = ConsoleColor.Black;
            Console.ForegroundColor = ConsoleColor.Cyan;
            Console.Write(message);
            LogLine(line);
            Console.ResetColor();
        }

        public static void LogLine(int lines = 1)
        {
            for (int i = 0; i < lines; i++)
            {
                Console.WriteLine();
            }
        }

        public static void LogSameLine(string message, int line = 0)
        {
            Console.BackgroundColor = ConsoleColor.Black;
            Console.ForegroundColor = ConsoleColor.White;
            Console.Write(message);
            LogLine(line);
            Console.ResetColor();
        }

        public static void LogWarning(string message, int line = 0)
        {
            Console.ForegroundColor = ConsoleColor.Yellow;
            LogLine(line);
            Console.WriteLine(message);
            Console.ResetColor();
        }

        public static void LogWarningAndExit(string message, int line = 1)
        {
            LogLine(line);
            LogWarning(message);
            ExitApplication();
        }

        private static void LogError(string message, int line = 0)
        {
            Console.ForegroundColor = ConsoleColor.Red;
            LogLine(line);
            Console.WriteLine(message);
            Console.ResetColor();
        }
    }
}