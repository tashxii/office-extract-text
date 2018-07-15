using OfficeExtractText.Properties;
using System;

namespace OfficeExtractText
{
    class ConsoleLogger
    {
        private static ConsoleColor defaultColor = Console.ForegroundColor;
        private static bool setToConsoleError = false;
        private static bool isEnableLog = true;
        private static bool isEnableWarning = true;
        private static bool isEnableError = true;

        internal static void InitLogger(bool isEnableLog, bool isEnableWarning, bool isEnableError)
        {
            ConsoleLogger.isEnableLog = isEnableLog;
            ConsoleLogger.isEnableWarning = isEnableWarning;
            ConsoleLogger.isEnableError = isEnableError;
        }

        internal static void WriteLog(string message)
        {
            if (isEnableLog)
            {
                Console.ForegroundColor = ConsoleColor.Cyan;
                Console.WriteLine(GetTimestampedMessage(message, Resource.LOG_PREFIX));
                Console.ForegroundColor = defaultColor;
            }
        }

        internal static void WriteWarning(string message)
        {
            if (isEnableWarning)
            {
                Console.ForegroundColor = ConsoleColor.Yellow;
                Console.WriteLine(GetTimestampedMessage(message, Resource.WARNING_PREFIX));
                Console.ForegroundColor = defaultColor;
            }
        }

        internal static void WriteError(string message)
        {
            if (isEnableError)
            {
                if (!setToConsoleError)
                {
                    ConsoleErrorWriter.SetToConsoleError();
                    setToConsoleError = true;
                }
                Console.Error.WriteLine(GetTimestampedMessage(message, Resource.ERROR_PREFIX));
            }
        }

        private static string GetTimestampedMessage(string message, string prefix)
        {
            DateTimeOffset now = DateTimeOffset.Now;
            return String.Format("{0}: {1} {2}", now.ToString("yyyy-MM-dd HH:mm:ss.fff"), prefix, message);
        }
    }
}
