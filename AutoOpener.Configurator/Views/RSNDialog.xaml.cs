using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using System.IO;

namespace AutoOpener.Configurator.Views
{
    /// <summary>
    /// Логика взаимодействия для RSNDialog.xaml
    /// </summary>
    public partial class RSNDialog : Window
    {
        public string RsnPath { get; set; }

        public RSNDialog(string initialHost = null, string initialProject = null, string initialModel = null)
        {
            InitializeComponent();
            tbHost.Text = initialHost ?? "";
            tbProject.Text = initialProject ?? "";
            tbModel.Text = initialModel ?? "";
            UpdatePreview();
            tbHost.TextChanged += (_, __) => UpdatePreview();
            tbProject.TextChanged += (_, __) => UpdatePreview();
            tbModel.TextChanged += (_, __) => UpdatePreview();
        }

        private void UpdatePreview()
        {
            var host = (tbHost.Text ?? "").Trim();
            var project = SanitizeSegment(tbProject.Text);
            var model = SanitizeModel(tbModel.Text);
            if (string.IsNullOrEmpty(host) || string.IsNullOrEmpty(project) || string.IsNullOrEmpty(model))
            {
                tbPreview.Text = "";
                return;
            }
            tbPreview.Text = $"RSN://{host}/{project}/{model}";
        }

        private static string SanitizeSegment(string s)
        {
            s = (s ?? "").Trim().Trim('/');
            foreach (var ch in System.IO.Path.GetInvalidFileNameChars())
                s = s.Replace(ch.ToString(), "");
            return s;
        }

        private static string SanitizeModel(string s)
        {
            s = SanitizeSegment(s);
            if (s.Length == 0) return s;
            if (!s.EndsWith(".rvt", StringComparison.OrdinalIgnoreCase))
                s += ".rvt";
            return s;
        }

        private void OnOk(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(tbPreview.Text))
            {
                MessageBox.Show(this, "Please fill all fields (Server, Project, Model).",
                    "AutoOpener", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }
            RsnPath = tbPreview.Text;
            DialogResult = true;
            Close();
        }
    }
}
