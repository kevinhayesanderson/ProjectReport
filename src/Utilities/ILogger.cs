namespace Utilities
{
    public interface ILogger
    {
        void Log(string message, int line = 0);

        void LogChar(char character, int length);

        void LogData(string data, int line = 0);

        void LogDataSameLine(string data, int line = 0);

        void LogError(string message, int line = 0);

        void LogInfo(string message, int line = 0);

        void LogInfoSameLine(string message, int line = 0);

        void LogLine(int lines = 1);

        void LogSameLine(string message, int line = 0);

        void LogWarning(string message, int line = 0);
    }
}