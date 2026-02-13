using AutoOpener.Core.IO;
using System.Windows;

namespace AutoOpener.Configurator
{
    /// <summary>
    /// Логика взаимодействия для App.xaml
    /// </summary>
    public partial class App : Application
    {
        protected override void OnStartup(StartupEventArgs e)
        {
            System.IO.Directory.CreateDirectory(PathsService.LogsDir);
            Logger.Info("Configurator start");
            base.OnStartup(e);
        }

        protected override void OnExit(ExitEventArgs e)
        {
            Logger.Info("Configurator exit");
            base.OnExit(e);
        }
    }
}
