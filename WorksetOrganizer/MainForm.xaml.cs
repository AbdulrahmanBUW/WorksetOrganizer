using Autodesk.Revit.UI;
using System;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Forms;

namespace WorksetOrchestrator
{
    public partial class MainForm : Window
    {
        private UIDocument _uiDoc;
        private WorksetOrchestrator _orchestrator;

        public MainForm(UIDocument uiDoc)
        {
            InitializeComponent();
            _uiDoc = uiDoc;
            _orchestrator = new WorksetOrchestrator(uiDoc);
            _orchestrator.LogUpdated += OnLogUpdated;

            // Set Revit as owner window
            this.Owner = RevitWindow.GetRevitWindow();
        }

        private void OnLogUpdated(object sender, string message)
        {
            // Thread-safe update of the log textbox
            Dispatcher.Invoke(() => {
                txtLog.AppendText(message + Environment.NewLine);
                txtLog.ScrollToEnd();
            });
        }

        private void BtnBrowseExcel_Click(object sender, RoutedEventArgs e)
        {
            var openFileDialog = new Microsoft.Win32.OpenFileDialog
            {
                Filter = "Excel Files (*.xlsx)|*.xlsx|All Files (*.*)|*.*",
                Title = "Select Workset Mapping Excel File"
            };

            if (openFileDialog.ShowDialog() == true)
            {
                txtExcelPath.Text = openFileDialog.FileName;
            }
        }

        private void BtnBrowseDestination_Click(object sender, RoutedEventArgs e)
        {
            var folderDialog = new FolderBrowserDialog
            {
                Description = "Select Destination Folder for Exports"
            };

            if (folderDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                txtDestination.Text = folderDialog.SelectedPath;
            }
        }

        private async void BtnRun_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(txtExcelPath.Text) || !File.Exists(txtExcelPath.Text))
            {
                System.Windows.MessageBox.Show("Please select a valid Excel file.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            if (string.IsNullOrEmpty(txtDestination.Text) || !Directory.Exists(txtDestination.Text))
            {
                System.Windows.MessageBox.Show("Please select a valid destination folder.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            btnRun.IsEnabled = false;
            txtLog.Clear();

            try
            {
                // Read mapping from Excel
                LogMessage("Reading Excel mapping file...");
                var mapping = ExcelReader.ReadMapping(txtExcelPath.Text);
                LogMessage($"Found {mapping.Count} mapping records.");

                // Process the mapping
                bool success = await System.Threading.Tasks.Task.Run(() =>
                    _orchestrator.Execute(mapping, txtDestination.Text, chkOverwrite.IsChecked == true, chkExportQc.IsChecked == true)
                );

                if (success)
                {
                    LogMessage("Process completed successfully!");
                    System.Windows.MessageBox.Show("Workset orchestration completed successfully!", "Success", MessageBoxButton.OK, MessageBoxImage.Information);
                }
                else
                {
                    LogMessage("Process completed with errors.");
                    System.Windows.MessageBox.Show("Workset orchestration completed with errors. Check the log for details.", "Warning", MessageBoxButton.OK, MessageBoxImage.Warning);
                }
            }
            catch (Exception ex)
            {
                LogMessage($"ERROR: {ex.Message}");
                System.Windows.MessageBox.Show($"An error occurred: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                btnRun.IsEnabled = true;
            }
        }

        private void BtnCancel_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void LogMessage(string message)
        {
            txtLog.AppendText($"{DateTime.Now:HH:mm:ss} - {message}{Environment.NewLine}");
            txtLog.ScrollToEnd();
        }
    }

    // Helper class to get Revit's main window handle
    public static class RevitWindow
    {
        public static Window GetRevitWindow()
        {
            return System.Windows.Application.Current.Windows.OfType<Window>().FirstOrDefault(w => w.Title.Contains("Revit"));
        }
    }
}