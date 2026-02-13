using System;
using System.IO;
using System.Linq;
using System.Windows;
using Microsoft.Win32;
using AutoOpener.Core.IO;
using AutoOpener.Core.Models;
using System.Collections.Generic;
using System.Runtime.Serialization.Json;

namespace AutoOpener.Configurator.Views
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            cbVersion.SelectedIndex = 1;
        }

        private void OnBrowse(object sender, RoutedEventArgs e)
        {
            var dlg = new OpenFileDialog
            {
                Filter = "Revit files (*.rvt)|*.rvt|All files (*.*)|*.*",
                CheckFileExists = true
            };
            if (dlg.ShowDialog(this) == true)
                tbModel.Text = dlg.FileName;
        }

        private int GetSelectrdVersion()
        {
            var item = cbVersion.SelectedItem as System.Windows.Controls.ComboBoxItem;
            if (item == null) return 2023;
            return int.TryParse(item.Content as string, out int v) ? v : 2023;
        }

        private void OnSaveProfile(object sender, RoutedEventArgs e)
        {
            try
            {
                var name = (tbName.Text ?? "").Trim();
                if (string.IsNullOrEmpty(name))
                {
                    MessageBox.Show(this, "Profile name is required.", "AutoOpener", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }
                var version = GetSelectrdVersion();
                var model = (tbModel.Text ?? "").Trim();
                if (string.IsNullOrEmpty(model))
                {
                    MessageBox.Show(this, "Model path is required.", "AutoOpener", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                var ws = (tbWorksets.Text ?? "").Trim();
                var list = ws.Length == 0 
                    ? new List<string>() 
                    : ws.Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries)
                    .Select(s => s.Trim()).Where(s => s.Length > 0).ToList();

                var profile = new Profile
                {
                    Name = name,
                    RevitVersion = version,
                    ModelPathOrRsn = model,
                    WorksetsByName = list,
                    OpenReadOnly = false
                };

                // ensure profiles dir
                var profilesDir = AutoOpener.Core.IO.PathsService.ProfileDir;
                var file = Path.Combine(profilesDir, SanitizeFileName(name) + ".json");

                // сохраняем JSON (через ваш JsonStorage, если есть метод Write<T>)
                try
                {
                    JsonStorage.Write(file, profile);
                }
                catch
                {
                    using (var fs = File.Create(file))
                    {
                        var ser = new DataContractJsonSerializer(typeof(Profile));
                        ser.WriteObject(fs, profile);
                    }
                }

                // если RSN — сразу добавим host в RSN.ini
                var host = RsnIniService.TryParseRsnHost(model);
                if (!string.IsNullOrEmpty(host))
                {
                    string rsnPath, action;
                    var ok = RsnIniService.EnsureServerListed(version, host, out rsnPath, out action);
                    Logger.Info($"Ensure RSN from SaveProfile: ok={ok}, action={action}, file={rsnPath ?? "(null)"}");
                }

                Logger.Info($"Profile saved: {file}");
                MessageBox.Show(this, "Profile saved.", "AutoOpener", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                Logger.Error("Save profile failed: " + ex);
                MessageBox.Show(this, "Save failed:\n" + ex.Message, "AutoOpener", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void OnEnsureRsn(object sender, RoutedEventArgs e)
        {
            var model = (tbModel.Text ?? "").Trim();
            var host = RsnIniService.TryParseRsnHost(model);
            if (string.IsNullOrEmpty(host))
            {
                MessageBox.Show(this, "Model path is not RSN://...", "AutoOpener",
                    MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            var version = GetSelectrdVersion();
            string rsnPath, action;
            var ok = RsnIniService.EnsureServerListed(version, host, out rsnPath, out action);

            if (ok)
            {
                MessageBox.Show(this,
                    $"RSN host '{host}' {action} for Revit {version}.\nFile: {rsnPath}",
                    "AutoOpener", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            else
            {
                MessageBox.Show(this,
                    $"Failed to ensure RSN host '{host}' for Revit {version}.\nAction: {action}\nFile: {rsnPath ?? "(unknown)"}\n\n" +
                    "Tip: try running Configurator as Administrator.",
                    "AutoOpener", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void OnClose(object sender, RoutedEventArgs e) => Close();

        private void OnBuildRsn(object sender, RoutedEventArgs e)
        {
            // Попытка распарсить текущий RSN для предзаполнения
            string host = null, project = null, model = null;
            var current = (tbModel.Text ?? "").Trim();
            if (current.StartsWith("RSN://", StringComparison.OrdinalIgnoreCase))
            {
                try
                {
                    var rest = current.Substring("RSN://".Length);
                    var parts = rest.Split(new[] { '/' }, StringSplitOptions.RemoveEmptyEntries);
                    if (parts.Length >= 1) host = parts[0];
                    if (parts.Length >= 2) project = parts[1];
                    if (parts.Length >= 3) model = parts[2];
                }
                catch { }
            }

            var dlg = new RSNDialog(host, project, model) { Owner = this };
            if (dlg.ShowDialog() == true && !string.IsNullOrEmpty(dlg.RsnPath))
            {
                tbModel.Text = dlg.RsnPath;
            }
        }

        private static string SanitizeFileName(string name)
        {
            foreach (var ch in Path.GetInvalidFileNameChars())
                name = name.Replace(ch, '_');
            return name;
        }
    }
}
