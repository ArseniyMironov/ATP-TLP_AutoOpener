using System;
using System.IO;

namespace AutoOpener.Core.IO
{
    public static class PathsService
    {
        // Базовый корень без версии: %AppData%\AutoOpener
        private static string RootBase =>
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "AutoOpener");

        // Контекст версии (задаётся add-in при старте). null = без версии.
        private static int? _versionContext;

        /// <summary>Установить контекст версии (например, 2022/2023) для текущего процесса.</summary>
        public static void SetVersionContext(int version) => _versionContext = version;

        /// <summary>Текущий корень с учётом контекста: %AppData%\AutoOpener\{year} или без года.</summary>
        public static string Root => _versionContext.HasValue
            ? Path.Combine(RootBase, _versionContext.Value.ToString())
            : RootBase;

        // ==== НЕверсионные (общие) каталоги ====
        // Профили оставляем общими (однажды созданные профили могут запускаться под разными версиями).
        public static string ProfileDir => Ensure(Path.Combine(RootBase, "profiles"));

        // ==== Версионные каталоги (Root зависит от SetVersionContext или можно вызывать *For(version)) ====
        public static string QueueDir => Ensure(Path.Combine(Root, "queue"));
        public static string LogsDir => Ensure(Path.Combine(Root, "logs"));
        public static string OutDir => Ensure(Path.Combine(Root, "out"));

        // Явные методы "для версии", если не хотим менять контекст процесса.
        public static string RootFor(int version) => Ensure(Path.Combine(RootBase, version.ToString()));
        public static string QueueDirFor(int version) => Ensure(Path.Combine(RootFor(version), "queue"));
        public static string LogsDirFor(int version) => Ensure(Path.Combine(RootFor(version), "logs"));
        public static string OutDirFor(int version) => Ensure(Path.Combine(RootFor(version), "out"));

        private static string Ensure(string path)
        {
            if (!Directory.Exists(path)) Directory.CreateDirectory(path);
            return path;
        }
    }
}
