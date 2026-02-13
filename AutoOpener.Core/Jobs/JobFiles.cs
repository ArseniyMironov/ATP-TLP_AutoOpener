using AutoOpener.Core.IO;
using System;
using System.IO;

namespace AutoOpener.Core.Jobs
{
    public static   class JobFiles
    {
        /// <summary>
        /// Перемещает job *.json → *.running. Возвращает новый путь.
        /// </summary>
        public static string MarkRunning(string jobJsonPath)
        {
            if (string.IsNullOrEmpty(jobJsonPath) || !File.Exists(jobJsonPath)) return null;
            var running = Path.ChangeExtension(jobJsonPath, ".running");
            TryMove(jobJsonPath, running);
            return File.Exists(running) ? running : null;
        }

        public static void MarkDone(string runningPath, JobResult result)
        {
            WriteResult(result);
            var done = Path.ChangeExtension(runningPath, ".done");
            TryMove(runningPath, done);
        }

        /// <summary>
        /// Помечает running как fail и пишет result в out/{jobId}.result.json.
        /// </summary>
        public static void MarkFail(string runningPath, JobResult result)
        {
            WriteResult(result);
            var fail = Path.ChangeExtension(runningPath, ".fail");
            TryMove(runningPath, fail);
        }

        /// <summary>
        /// Путь к result-файлу по JobId.
        /// </summary>
        public static string GetResultPath(Guid jobId)
        {
            return Path.Combine(PathsService.OutDir, jobId.ToString("N") + ".result.json");
        }

        private static void WriteResult(JobResult result)
        {
            try
            {
                if (result == null || result.JobId == Guid.Empty) return;
                Directory.CreateDirectory(PathsService.OutDir);
                var file = GetResultPath(result.JobId);
                result.CompletedUtc = DateTime.UtcNow;
                JsonStorage.Write(file, result);
            }
            catch
            {

            }
        }

        private static void TryMove(string from, string to)
        {
            try
            {
                if (File.Exists(from))
                {
                    if (File.Exists(to)) File.Delete(to);
                    File.Move(from, to);
                }
            }
            catch { }
        }
    }
}
