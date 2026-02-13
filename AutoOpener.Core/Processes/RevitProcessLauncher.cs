using Microsoft.Win32;
using System;
using System.Diagnostics;
using System.IO;
using System.Security.Cryptography;
using System.Text;

namespace AutoOpener.Core.Processes
{
    public static class RevitProcessLauncher
    {
        /// <summary>
        /// Старый API для обратной совместимости (smoke и т.п.).
        /// Всегда просто стартует Revit указанной версии.
        /// </summary>
        public static bool TryStart(int version)
        {
            var exe = FindRevitExe(version);
            if (exe == null || !File.Exists(exe)) return false;
            Process.Start(new ProcessStartInfo(exe) { UseShellExecute = true });
            return true;
        }

        /// <summary>
        /// Новый API: перед стартом новой сессии проверяет, не открыта ли уже нужная модель (по lock-файлу).
        /// Если модель уже открыта в любой сессии Revit этой версии — НЕ запускаем новую сессию и возвращаем true.
        /// Иначе — запускаем новую сессию Revit {version}.
        /// ВАЖНО: лончер lock-файлы НЕ удаляет (уборку сделаем в add-in строго для своей сессии).
        /// </summary>
        public static bool TryStart(int version, string modelPathOrRsn)
        {
            if (!string.IsNullOrWhiteSpace(modelPathOrRsn) &&
                IsModelOpenViaLock(version, modelPathOrRsn))
            {
                // Модель уже открыта где-то -> не стартуем новый Revit
                return true;
            }

            var exe = FindRevitExe(version);
            if (exe == null || !File.Exists(exe)) return false;

            Process.Start(new ProcessStartInfo(exe) { UseShellExecute = true });
            return true;
        }

        /// <summary>
        /// Возвращает полный путь к Revit.exe для версии (или null).
        /// </summary>
        public static string TryGetExePath(int year) => FindRevitExe(year);

        /// <summary>
        /// Есть ли уже запущенный Revit нужной версии (сопоставляем по пути к exe).
        /// </summary>
        public static bool IsRevitRunning(int year)
        {
            var exe = FindRevitExe(year);
            if (string.IsNullOrEmpty(exe)) return false;

            var target = exe.ToLowerInvariant();
            foreach (var p in Process.GetProcessesByName("Revit"))
            {
                try
                {
                    var path = p.MainModule?.FileName;
                    if (!string.IsNullOrEmpty(path) && path.ToLowerInvariant() == target)
                        return true;
                }
                catch
                {
                    // Нет доступа к MainModule — пропускаем процесс
                }
            }
            return false;
        }

        // --------- LOCK: проверка "модель уже открыта" ---------

        /// <summary>
        /// Путь к lock-файлу модели: %AppData%\AutoOpener\locks\{year}\{md5(key)}.lock
        /// </summary>
        public static string GetModelLockPath(int year, string modelPathOrRsn)
        {
            var key = NormalizeModelKey(modelPathOrRsn);
            var hash = ComputeHash(key);
            var dir = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                "AutoOpener", "locks", year.ToString());
            return Path.Combine(dir, hash + ".lock");
        }

        /// <summary>
        /// True, если по lock-файлу видно, что модель уже открыта живым процессом Revit.
        /// Формат lock: "PID|VISIBLE_PATH|UTC_ISO8601"
        /// </summary>
        public static bool IsModelOpenViaLock(int year, string modelPathOrRsn)
        {
            try
            {
                var lf = GetModelLockPath(year, modelPathOrRsn);
                if (!File.Exists(lf)) return false;

                var text = File.ReadAllText(lf).Trim();
                if (string.IsNullOrEmpty(text)) return false;

                var parts = text.Split('|');
                if (parts.Length == 0) return false;

                int pid;
                if (!int.TryParse(parts[0], out pid)) return false;

                try
                {
                    var proc = Process.GetProcessById(pid);
                    if (proc != null && !proc.HasExited &&
                        string.Equals(proc.ProcessName, "Revit", StringComparison.OrdinalIgnoreCase))
                    {
                        // Живой процесс Revit — считаем модель открытой
                        return true;
                    }
                }
                catch
                {
                    // Процесс не жив/недоступен — НЕ удаляем lock здесь
                    return false;
                }
            }
            catch
            {
                // Любая ошибка интерпретации lock — лучше вернуть "не открыт"
            }
            return false;
        }

        private static string NormalizeModelKey(string path)
        {
            if (string.IsNullOrWhiteSpace(path)) return string.Empty;
            var s = path.Trim();
            s = s.Replace('\\', '/');          // слеши в единый формат
            s = s.ToUpperInvariant();          // безрегистровое сравнение
            // Доп. нормализация RSN тут не обязательна — ключ детерминирован.
            return s;
        }

        private static string ComputeHash(string s)
        {
            if (s == null) s = string.Empty;
            using (var md5 = MD5.Create())
            {
                var bytes = Encoding.UTF8.GetBytes(s);
                var hash = md5.ComputeHash(bytes);
                var sb = new StringBuilder(hash.Length * 2);
                foreach (var b in hash) sb.Append(b.ToString("x2"));
                return sb.ToString();
            }
        }

        // --------- Поиск Revit.exe ---------

        private static string FindRevitExe(int year)
        {
            // 1) Прямой InstallLocation (64 → 32)
            var exe = TryFromInstallLocation(year, RegistryView.Registry64)
                   ?? TryFromInstallLocation(year, RegistryView.Registry32);
            if (File.Exists(exe)) return exe;

            // 2) Через ProductCode → Uninstall (64 → 32)
            exe = TryFromUninstallProductCode(year, RegistryView.Registry64)
               ?? TryFromUninstallProductCode(year, RegistryView.Registry32);
            if (File.Exists(exe)) return exe;

            // 3) Дефолтная папка
            var fallback = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles),
                "Autodesk", $"Revit {year}", "Revit.exe");
            return File.Exists(fallback) ? fallback : null;
        }

        private static string TryFromInstallLocation(int year, RegistryView view)
        {
            try
            {
                using (var hklm = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, view))
                using (var key = hklm.OpenSubKey($@"SOFTWARE\Autodesk\Revit\Autodesk Revit {year}"))
                {
                    var install = key?.GetValue("InstallLocation") as string;
                    if (!string.IsNullOrEmpty(install))
                    {
                        var exe = Path.Combine(install, "Revit.exe");
                        return exe;
                    }
                }
            }
            catch { }
            return null;
        }

        private static string TryFromUninstallProductCode(int year, RegistryView view)
        {
            try
            {
                using (var hklm = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, view))
                using (var comp = hklm.OpenSubKey($@"SOFTWARE\Autodesk\Revit\Autodesk Revit {year}\Components"))
                {
                    var productCode = comp?.GetValue("ProductCode") as string;
                    if (string.IsNullOrEmpty(productCode)) return null;

                    using (var uninstall = hklm.OpenSubKey($@"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{productCode}"))
                    {
                        var install = uninstall?.GetValue("InstallLocation") as string;
                        if (!string.IsNullOrEmpty(install))
                        {
                            var exe = Path.Combine(install, "Revit.exe");
                            return exe;
                        }
                    }
                }
            }
            catch { }
            return null;
        }
    }
}