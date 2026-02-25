using AutoOpener.Core.IO;
using AutoOpener.Core.Jobs;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Forms;
using Application = System.Windows.Application;

namespace AutoOpener.Tray
{
    /// <summary>
    /// Логика взаимодействия для App.xaml
    /// </summary>
    public partial class App : Application
    {
        private NotifyIcon _notifyIcon;
        private List<FileSystemWatcher> _watchers = new List<FileSystemWatcher>();
        private static Mutex _mutex;

        protected override void OnStartup(StartupEventArgs e)
        {
            // Явно говорим WPF не закрываться, пока мы сами не вызовем Shutdown()
            this.ShutdownMode = ShutdownMode.OnExplicitShutdown;

            // Инициализация мьютекса
            const string appName = "AutoOpenerTray_SingleInstance_Mutex";
            _mutex = new Mutex(true, appName, out bool createdNew);

            if (!createdNew)
            {
                Application.Current.Shutdown();
                return;
            }

            base.OnStartup(e);

            // 1. Инициализация иконки в системном трее
            _notifyIcon = new NotifyIcon
            {
                Icon = SystemIcons.Information, // Заменить на свою иконку 
                Visible = true,
                Text = "AutoOpener Agent"
            };

            // Добавление контекстного меню для закрытия агента
            var menu = new ContextMenu();
            menu.MenuItems.Add("Выход", (s, ev) => Shutdown());
            _notifyIcon.ContextMenu = menu;

            // 2. Настройка слежения за папками результатов для обеих версий
            int[] versions = { 2022, 2023 };
            foreach (var v in versions)
            {
                try
                {
                    var outDir = PathsService.OutDirFor(v);
                    Directory.CreateDirectory(outDir);

                    var watcher = new FileSystemWatcher(outDir, "*.result.json");
                    watcher.Created += OnResultFileCreated;
                    watcher.EnableRaisingEvents = true;
                    _watchers.Add(watcher);
                }
                catch (Exception ex)
                {
                    Logger.Error($"[TRAY] Failed to start watcher for {v}: {ex.Message}");
                }
            }
        }

        private void OnResultFileCreated(object sender, FileSystemEventArgs e)
        {
            // Пауза, чтобы пишущий процесс (Revit) успел освободить файл
            System.Threading.Thread.Sleep(500);

            try
            {
                var result = JsonStorage.Read<JobResult>(e.FullPath);
                if (result != null)
                {
                    string modelName = string.IsNullOrEmpty(result.ModelPath) 
                        ? "Неизвестная модель" 
                        : Path.GetFileNameWithoutExtension(result.ModelPath);

                    string title = result.Succeeded ? 
                        $"Revit {result.RevitVersion}: Успех \n {modelName} открыта" 
                        : $"Revit {result.RevitVersion}: Ошибка \n Ошибька при открытии {modelName}";

                    ToolTipIcon iconType = result.Succeeded ? ToolTipIcon.Info : ToolTipIcon.Error;

                    // Вызов UI-компонента должен происходить в основном потоке
                    Dispatcher.Invoke(() =>
                    {
                        _notifyIcon.ShowBalloonTip(5000, title, result.Message, iconType);
                    });
                }
            }
            catch (Exception ex)
            {
                Logger.Error($"[TRAY] Failed to read result {e.FullPath}: {ex.Message}");
            }
        }

        protected override void OnExit(ExitEventArgs e)
        {
            // Корректное освобождение ресурсов при выходе
            if (_notifyIcon != null)
            {
                _notifyIcon.Visible = false;
                _notifyIcon.Dispose();
            }

            foreach (var w in _watchers)
            {
                w.EnableRaisingEvents = false;
                w.Dispose();
            }

            // осовбождаем mutex
            if (_mutex != null)
            {
                _mutex.ReleaseMutex();
                _mutex.Dispose();
            }

            base.OnExit(e);

            // Гарантированно убиваем процесс вместе со всеми фоновыми потоками ОС
            Environment.Exit(0);
        }
    }
}
