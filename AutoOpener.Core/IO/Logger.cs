using System;
using System.IO;

namespace AutoOpener.Core.IO
{
    public static class Logger
    {
        private static string FilePath =>
            Path.Combine(PathsService.LogsDir, DateTime.UtcNow.ToString("yyyyMMdd") + ".log");

        public static void Info(string msg) => Write("INFO", msg);
        public static void Error(string msg) => Write("ERROR", msg);

        private static void Write(string level, string msg)
        {
            try
            {
                Directory.CreateDirectory(PathsService.LogsDir);
                File.AppendAllText(FilePath, $"{DateTime.UtcNow:O} [{level}] {msg}{Environment.NewLine}");
            }
            catch
            {
                // намеренно глушим — лог не должен ломать процесс
            }
        }
        static Logger()
        {
            try
            {
                Directory.CreateDirectory(PathsService.LogsDir);
                File.AppendAllText(FilePath, DateTime.UtcNow.ToString("O") + " [BOOT] Logger ready\r");
            }
            catch { }
        }
    }
}
