using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AutoOpener.Core.IO
{
    public static class RsnIniService
    {
        /// <summary>
        /// Основной метод: пытается создать/обновить RSN.ini.
        /// Возвращает true при успехе (добавлен или уже присутствует), false при ошибке.
        /// action: "created and added" | "added" | "already present" | "failed: <reason>"
        /// rsnFile: полный путь к RSN.ini
        /// </summary>
        public static bool EnsureServerListed(int year, string hostName, out string rsnFile, out string action)
        {
            rsnFile = null;
            action = null;

            if (string.IsNullOrWhiteSpace(hostName))
            {
                action = "failed:empty";
                return false;
            }

            try
            {
                string baseDir = Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData);
                string rsnDir = Path.Combine(baseDir, "Autodesk", $"Revit Server {year}", "Config");
                rsnFile = Path.Combine(rsnDir, "RSN.ini");

                Directory.CreateDirectory(rsnDir);

                if (!File.Exists(rsnFile))
                {
                    File.WriteAllText(rsnFile, hostName + Environment.NewLine);
                    Logger.Info($"RSN.ini created ({year}), added host: {hostName}");
                    action = "created and added";
                    return true;
                }

                var lines = File.ReadAllLines(rsnFile)
                                .Select(l => (l ?? string.Empty).Trim())
                                .Where(l => !string.IsNullOrEmpty(l))
                                .ToList();

                bool exists = lines.Any(l => string.Equals(l, hostName, StringComparison.OrdinalIgnoreCase));
                if (exists)
                {
                    Logger.Info($"RSN.ini ({year}) already contains host '{hostName}'.");
                    action = "already present";
                    return true;
                }

                File.AppendAllText(rsnFile, hostName + Environment.NewLine);
                Logger.Info($"RSN.ini ({year}) updated: added host '{hostName}'.");
                action = "added";
                return true;
            }
            catch (UnauthorizedAccessException uae)
            {
                Logger.Error($"RSN.ini update failed ({year}): access denied. {uae.Message}");
                action = "failed: access denied";
                return false;
            }
            catch (Exception ex)
            {
                Logger.Error($"RSN.ini update failed ({year}): {ex}");
                action = "failed: " + ex.Message;
                return false;
            }
        }

        /// <summary>
        /// Быстрый парсер RSN://host/... → host
        /// </summary>
        public static string TryParseRsnHost(string rsnPath)
        {
            if (string.IsNullOrWhiteSpace(rsnPath)) return null;
            const string prefix = "RSN://";
            if (!rsnPath.StartsWith(prefix, StringComparison.OrdinalIgnoreCase)) return null;

            string rest = rsnPath.Substring(prefix.Length);
            int slash = rest.IndexOf('/');
            if (slash < 0) return rest.Trim();
            return rest.Substring(0, slash).Trim();
        }
    }
}
