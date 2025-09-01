using Autodesk.Revit.UI;
using System;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Interop;
using System.Windows.Forms; // Only for FolderBrowserDialog

namespace WorksetOrchestrator
{
    public partial class MainForm : Window
    {
        private UIDocument _uiDoc;
        private WorksetOrchestrator _orchestrator;

        public MainForm(UIDocument uiDoc, IntPtr revitMainWindowHandle)
        {
            InitializeComponent();
            _uiDoc = uiDoc;
            _orchestrator = new WorksetOrchestrator(uiDoc);
            _orchestrator.LogUpdated += OnLogUpdated;

            // Set Revit as owner window via handle if available
            try
            {
                if (revitMainWindowHandle != IntPtr.Zero)
                {
                    var helper = new WindowInteropHelper(this);
                    helper.Owner = revitMainWindowHandle;
                }
                else
                {
                    var revitWindow = RevitWindow.GetRevitWindow();
                    if (revitWindow != null)
                        this.Owner = revitWindow;
                }
            }
            catch
            {
                // Owner window is optional
            }
        }

        private void OnLogUpdated(object sender, string message)
        {
            // Thread-safe log update
            Dispatcher.Invoke(() =>
            {
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

            bool? result = openFileDialog.ShowDialog(); // WPF OpenFileDialog returns bool?
            if (result == true)
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

        private void BtnRun_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(txtExcelPath.Text) || !File.Exists(txtExcelPath.Text))
            {
                System.Windows.MessageBox.Show("Please select a valid Excel file.", "Error",
                    MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            if (string.IsNullOrEmpty(txtDestination.Text) || !Directory.Exists(txtDestination.Text))
            {
                System.Windows.MessageBox.Show("Please select a valid destination folder.", "Error",
                    MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            btnRun.IsEnabled = false;
            txtLog.Clear();

            try
            {
                LogMessage("Reading Excel mapping file...");
                var mapping = ExcelReader.ReadMapping(txtExcelPath.Text);
                LogMessage($"Found {mapping.Count} mapping records.");

                // Execute orchestrator synchronously on Revit thread
                bool success = _orchestrator.Execute(mapping, txtDestination.Text,
                    chkOverwrite.IsChecked == true, chkExportQc.IsChecked == true);

                if (success)
                {
                    LogMessage("Process completed successfully!");
                    System.Windows.MessageBox.Show("Workset orchestration completed successfully!", "Success",
                        MessageBoxButton.OK, MessageBoxImage.Information);
                }
                else
                {
                    LogMessage("Process completed with errors.");
                    System.Windows.MessageBox.Show("Workset orchestration completed with errors. Check the log for details.",
                        "Warning", MessageBoxButton.OK, MessageBoxImage.Warning);
                }
            }
            catch (Exception ex)
            {
                LogMessage($"ERROR: {ex.Message}");
                System.Windows.MessageBox.Show($"An error occurred: {ex.Message}", "Error",
                    MessageBoxButton.OK, MessageBoxImage.Error);
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

    // Helper class to get Revit's main window handle safely
    public static class RevitWindow
    {
        public static System.Windows.Window GetRevitWindow()
        {
            try
            {
                var app = System.Windows.Application.Current;
                if (app == null) return null;

                return app.Windows
                          .OfType<System.Windows.Window>()
                          .FirstOrDefault(w => !string.IsNullOrEmpty(w?.Title) &&
                                               w.Title.IndexOf("Revit", StringComparison.OrdinalIgnoreCase) >= 0);
            }
            catch
            {
                return null;
            }
        }
    }
}
