using System;
using System.IO;
using System.Linq;

namespace AutoOpener.Core.IO
{
    public static class CleanupService
    {
        public static void CleanupOldArtifacts(int version, int retentionDays = 7)
        {
            var cutoffUtc = DateTime.UtcNow.AddDays(-Math.Max(1, retentionDays));

            // queue
            SafeDeleteByMask(PathsService.QueueDirFor(version), cutoffUtc, "*.done", "*.fail", "*.bad");

            // out
            SafeDeleteByMask(PathsService.QueueDirFor(version), cutoffUtc, "*.result.json");

            // logs
            SafeDeleteByMask(PathsService.QueueDirFor(version), cutoffUtc, "*.*");
        }

        private static void SafeDeleteByMask(string dir, DateTime cutoffUtc, params string[] masks)
        {
            try
            {
                if (string.IsNullOrEmpty(dir) || !Directory.Exists(dir)) return;
                foreach (var mask in masks)
                {
                    string[] files;
                    try { files = Directory.GetFiles(dir, mask); } catch { continue; }

                    foreach (var f in files)
                    {
                        try
                        {
                            var t = File.GetLastWriteTimeUtc(f);
                            if (t < cutoffUtc)
                                File.Delete(f);
                        }
                        catch { /* не роняем процесс */ }
                    }
                }
            }
            catch { /* глушим всё */ }
        }
    }
}
