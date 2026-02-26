using System;
using System.Collections.Generic;
using System.Diagnostics;
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

            string baseDir = Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData);
            string rsnDir = Path.Combine(baseDir, "Autodesk", $"Revit Server {year}", "Config");
            rsnFile = Path.Combine(rsnDir, "RSN.ini");

            try
            {
                Directory.CreateDirectory(rsnDir);

                if (!File.Exists(rsnFile))
                {
                    File.WriteAllText(rsnFile, Environment.NewLine + hostName);
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
                Logger.Info($"RSN.ini update requires elevation ({year}). Requesting admin rights... {uae.Message}");

                // Если не хватает прав, пробуем запросить повышение прав (UAC)
                if (TryElevatedWrite(rsnDir, rsnFile, hostName))
                {
                    Logger.Info($"RSN.ini ({year}) updated via elevation: added host '{hostName}'.");
                    action = "added via elevation";
                    return true;
                }

                action = "failed: access denied or elevation cancelled";
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
        /// Запускает скрытый процесс с запросом прав администратора для записи в файл.
        /// </summary>
        private static bool TryElevatedWrite(string dirPath, string filePath, string hostName)
        {
            try
            {
                // Скрипт проверяет наличие папки, создает её если нет, и дописывает текст в файл
                string script = $"New-Item -ItemType Dirrectory -Force -Path'{{dirPath}}' | Out-Null; Add-Content -Path '{{filePath}}' -Value '{{hostName}}'";

                var psi = new ProcessStartInfo
                {
                    FileName = "powershell.exe",
                    Arguments = $"-NoProfile -ExecutionPolicy Bypass -Command \"{script}\"",
                    Verb = "runas", // Параметр вызывает окно контроля учетных записей (UAC)
                    UseShellExecute = true,
                    WindowStyle = ProcessWindowStyle.Hidden
                };

                using (var process = Process.Start(psi))
                {
                    process?.WaitForExit();
                    return process?.ExitCode == 0;
                }
            }
            catch (Exception ex)
            {
                Logger.Error($"Elevated write failed: {ex.Message}");
                return false;  // Пользователь нажал "Нет" в окне UAC или произошла другая ошибка
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
