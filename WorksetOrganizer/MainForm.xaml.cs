using Autodesk.Revit.UI;
using System;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Interop;
using System.Windows.Forms; // Only for FolderBrowserDialog

namespace WorksetOrchestrator
{
    public partial class MainForm : Window
    {
        private UIDocument _uiDoc;
        private WorksetOrchestrator _orchestrator;
        private WorksetEventHandler _eventHandler;
        private ExternalEvent _externalEvent;

        public MainForm(UIDocument uiDoc, IntPtr revitMainWindowHandle)
        {
            InitializeComponent();
            _uiDoc = uiDoc;
            _orchestrator = new WorksetOrchestrator(uiDoc);
            _orchestrator.LogUpdated += OnLogUpdated;

            // Create external event handler for API calls
            _eventHandler = new WorksetEventHandler();
            _externalEvent = ExternalEvent.Create(_eventHandler);

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

            // Log initial information
            LogMessage($"Document: {_uiDoc.Document.Title}");
            LogMessage($"Workshared: {_uiDoc.Document.IsWorkshared}");
        }

        protected override void OnClosed(EventArgs e)
        {
            // Dispose of external event when window closes
            _externalEvent?.Dispose();
            base.OnClosed(e);
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

            bool? result = openFileDialog.ShowDialog();
            if (result == true)
            {
                txtExcelPath.Text = openFileDialog.FileName;
                LogMessage($"Selected Excel file: {System.IO.Path.GetFileName(openFileDialog.FileName)}");
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
                LogMessage($"Selected destination: {folderDialog.SelectedPath}");
            }
        }

        private async void BtnRun_Click(object sender, RoutedEventArgs e)
        {
            // Validation
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

            if (!_uiDoc.Document.IsWorkshared)
            {
                System.Windows.MessageBox.Show("The current document is not workshared. Please enable worksharing first.", "Error",
                    MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            btnRun.IsEnabled = false;
            btnCancel.Content = "Close";
            txtLog.Clear();

            try
            {
                LogMessage("=== WORKSET ORCHESTRATION STARTED ===");
                LogMessage("Reading Excel mapping file...");

                var mapping = ExcelReader.ReadMapping(txtExcelPath.Text);
                LogMessage($"Found {mapping.Count} mapping records.");

                // Log mapping summary
                foreach (var record in mapping.Take(5)) // Show first 5 for verification
                {
                    LogMessage($"  Pattern: '{record.SystemNameInModel}' → Workset: '{record.WorksetName}' → Package: '{record.ModelPackageCode}'");
                }
                if (mapping.Count > 5)
                    LogMessage($"  ... and {mapping.Count - 5} more records.");

                // Set parameters for the external event handler
                _eventHandler.SetParameters(_orchestrator, mapping, txtDestination.Text,
                    chkOverwrite.IsChecked == true, chkExportQc.IsChecked == true);

                // Raise the external event to execute in Revit context
                LogMessage("Starting workset orchestration in Revit context...");
                ExternalEventRequest request = _externalEvent.Raise();

                if (request == ExternalEventRequest.Accepted)
                {
                    // Wait for completion with timeout
                    await WaitForCompletionAsync();

                    if (_eventHandler.Success)
                    {
                        LogMessage("=== PROCESS COMPLETED SUCCESSFULLY ===");
                        System.Windows.MessageBox.Show("Workset orchestration completed successfully!", "Success",
                            MessageBoxButton.OK, MessageBoxImage.Information);
                    }
                    else
                    {
                        LogMessage("=== PROCESS COMPLETED WITH ERRORS ===");
                        if (_eventHandler.LastException != null)
                        {
                            LogMessage($"Last exception: {_eventHandler.LastException.Message}");
                        }
                        System.Windows.MessageBox.Show("Workset orchestration completed with errors. Check the log for details.",
                            "Warning", MessageBoxButton.OK, MessageBoxImage.Warning);
                    }
                }
                else
                {
                    LogMessage($"ERROR: Failed to raise external event. Request status: {request}");
                    System.Windows.MessageBox.Show("Failed to start the process. Make sure Revit is active and the document is not in edit mode.",
                        "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
            catch (Exception ex)
            {
                LogMessage($"ERROR: {ex.Message}");
                LogMessage($"Stack Trace: {ex.StackTrace}");
                System.Windows.MessageBox.Show($"An error occurred: {ex.Message}", "Error",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                btnRun.IsEnabled = true;
                btnCancel.Content = "Cancel";
            }
        }

        private async Task WaitForCompletionAsync()
        {
            // Wait for the external event to complete with a reasonable timeout
            int maxWaitTime = 600000; // 10 minutes for large models
            int checkInterval = 250; // 250ms
            int totalWaited = 0;
            int lastLogTime = 0;

            while (!_eventHandler.IsComplete && totalWaited < maxWaitTime)
            {
                await Task.Delay(checkInterval);
                totalWaited += checkInterval;

                // Log progress every 30 seconds
                if (totalWaited - lastLogTime >= 30000)
                {
                    LogMessage($"Processing... ({totalWaited / 1000} seconds elapsed)");
                    lastLogTime = totalWaited;
                }

                // Allow UI to update
                System.Windows.Forms.Application.DoEvents();
            }

            if (totalWaited >= maxWaitTime)
            {
                LogMessage("WARNING: Operation timed out after 10 minutes.");
            }
        }

        private void BtnCancel_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void LogMessage(string message)
        {
            string timestampedMessage = $"{DateTime.Now:HH:mm:ss} - {message}";
            txtLog.AppendText(timestampedMessage + Environment.NewLine);
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