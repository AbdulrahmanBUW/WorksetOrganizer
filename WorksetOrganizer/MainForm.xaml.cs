using Autodesk.Revit.UI;
using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Interop;
using System.Windows.Forms; // Only for FolderBrowserDialog
using System.Collections.Generic;

namespace WorksetOrchestrator
{
    public partial class MainForm : Window
    {
        private UIDocument _uiDoc;
        private WorksetOrchestrator _orchestrator;
        private WorksetEventHandler _eventHandler;
        private ExternalEvent _externalEvent;
        private List<string> _extractedFiles = new List<string>();
        private string _lastDestinationPath;
        private bool _isExtractWorksetMode = false;

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

            // Initialize UI
            txtStatus.Text = "Ready";
            pbProgress.Value = 0;
            pbProgress.IsIndeterminate = false;
            btnIntegrateTemplate.Visibility = Visibility.Collapsed; // Initially hidden

            // Log initial information
            try
            {
                LogMessage($"Document: {_uiDoc.Document.Title}");
                LogMessage($"Workshared: {_uiDoc.Document.IsWorkshared}");
            }
            catch (Exception ex)
            {
                LogMessage($"Warning reading document info: {ex.Message}");
            }

            // Initialize in QC mode
            SetQcMode();
        }

        protected override void OnClosed(EventArgs e)
        {
            // Dispose of external event when window closes
            _externalEvent?.Dispose();
            base.OnClosed(e);
        }

        private void SetQcMode()
        {
            _isExtractWorksetMode = false;
            txtModeDescription.Text = "QC Check & Extraction Mode";
            txtCurrentMode.Text = "QC Mode: Organizes and validates worksets based on Excel mapping";

            // Show Excel file selection and options
            gridExcelFile.Visibility = Visibility.Visible;
            stackPanelOptions.Visibility = Visibility.Visible;

            // Update button states
            btnQcCheck.Background = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(45, 137, 239)); // Primary color
            btnExtractWorksets.Background = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(240, 244, 250)); // Secondary color
            btnExtractWorksets.Foreground = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(51, 51, 51)); // Dark text

            LogMessage("Switched to QC Check & Extraction mode");
        }

        private void SetExtractWorksetsMode()
        {
            _isExtractWorksetMode = true;
            txtModeDescription.Text = "Extract Worksets Mode";
            txtCurrentMode.Text = "Extract Mode: Extracts all available worksets into separate files";

            // Hide Excel file selection and QC options (destination is still needed)
            gridExcelFile.Visibility = Visibility.Collapsed;
            stackPanelOptions.Visibility = Visibility.Collapsed;

            // Update button states
            btnExtractWorksets.Background = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(255, 140, 0)); // Orange color
            btnExtractWorksets.Foreground = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Colors.White);
            btnQcCheck.Background = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(240, 244, 250)); // Secondary color
            btnQcCheck.Foreground = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(51, 51, 51)); // Dark text

            LogMessage("Switched to Extract Worksets mode");
        }

        private void OnLogUpdated(object sender, string message)
        {
            // Thread-safe log update
            Dispatcher.Invoke(() =>
            {
                txtLog.AppendText(message + Environment.NewLine);
                txtLog.ScrollToEnd();

                // Keep the status line short and informative
                if (!string.IsNullOrEmpty(message))
                {
                    // show the last log short text in status (trim timestamp)
                    int hyphenIndex = message.IndexOf(" - ");
                    string shortMsg = message;
                    if (hyphenIndex >= 0 && hyphenIndex + 3 < message.Length)
                        shortMsg = message.Substring(hyphenIndex + 3);

                    txtStatus.Text = shortMsg.Length > 80 ? shortMsg.Substring(0, 77) + "..." : shortMsg;
                }
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

        private async void BtnQcCheck_Click(object sender, RoutedEventArgs e)
        {
            // Switch to QC mode if not already
            if (_isExtractWorksetMode)
            {
                SetQcMode();
                return;
            }

            // Run QC Check & Extraction
            await RunQcCheckAndExtraction();
        }

        private async void BtnExtractWorksets_Click(object sender, RoutedEventArgs e)
        {
            // Switch to Extract mode if not already
            if (!_isExtractWorksetMode)
            {
                SetExtractWorksetsMode();
                return;
            }

            // Run Extract Worksets
            await RunExtractWorksets();
        }

        private async Task RunQcCheckAndExtraction()
        {
            // Validation for QC mode
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

            await ExecuteOperation("QC Check & Extraction", async () =>
            {
                LogMessage("=== QC CHECK & EXTRACTION STARTED ===");
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
                _externalEvent.Raise();

                // Wait for completion with timeout
                await WaitForCompletionAsync();
            });
        }

        private async Task RunExtractWorksets()
        {
            // Validation for Extract mode
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

            await ExecuteOperation("Extract Worksets", async () =>
            {
                LogMessage("=== EXTRACT WORKSETS STARTED ===");
                LogMessage("Extracting all available worksets...");

                // Set parameters for workset extraction mode
                _eventHandler.SetWorksetExtractionParameters(_orchestrator, txtDestination.Text, true); // Always overwrite in extract mode

                // Raise the external event to execute in Revit context
                LogMessage("Starting workset extraction in Revit context...");
                _externalEvent.Raise();

                // Wait for completion with timeout
                await WaitForCompletionAsync();
            });
        }

        private async Task ExecuteOperation(string operationName, Func<Task> operation)
        {
            // Reset extracted files list and hide template button
            _extractedFiles.Clear();
            btnIntegrateTemplate.Visibility = Visibility.Collapsed;
            _lastDestinationPath = txtDestination.Text;

            // Disable UI while running
            btnQcCheck.IsEnabled = false;
            btnExtractWorksets.IsEnabled = false;
            btnCancel.Content = "Close";
            txtLog.Clear();

            // Show progress
            pbProgress.IsIndeterminate = true;
            txtStatus.Text = "Initializing...";

            try
            {
                await operation();

                // After the wait, update UI based on result
                pbProgress.IsIndeterminate = false;
                pbProgress.Value = 100;

                if (_eventHandler.Success)
                {
                    LogMessage($"=== {operationName.ToUpper()} COMPLETED SUCCESSFULLY ===");
                    txtStatus.Text = "Completed successfully";

                    // Get list of extracted files for template integration
                    _extractedFiles = GetExtractedFiles(txtDestination.Text);
                    LogMessage($"Found {_extractedFiles.Count} extracted files for potential template integration.");

                    // Show template integration button if there are extracted files
                    if (_extractedFiles.Count > 0)
                    {
                        btnIntegrateTemplate.Visibility = Visibility.Visible;
                        LogMessage("Click 'Integrate into Template' to copy extracted elements into a template file.");
                    }

                    System.Windows.MessageBox.Show($"{operationName} completed successfully!", "Success",
                        MessageBoxButton.OK, MessageBoxImage.Information);
                }
                else
                {
                    LogMessage($"=== {operationName.ToUpper()} COMPLETED WITH ERRORS ===");
                    txtStatus.Text = "Completed with errors";
                    if (_eventHandler.LastException != null)
                    {
                        LogMessage($"Last exception: {_eventHandler.LastException.Message}");
                    }
                    System.Windows.MessageBox.Show($"{operationName} completed with errors. Check the log for details.",
                        "Warning", MessageBoxButton.OK, MessageBoxImage.Warning);
                }
            }
            catch (Exception ex)
            {
                LogMessage($"ERROR: {ex.Message}");
                LogMessage($"Stack Trace: {ex.StackTrace}");
                txtStatus.Text = "Error";
                System.Windows.MessageBox.Show($"An error occurred: {ex.Message}", "Error",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                btnQcCheck.IsEnabled = true;
                btnExtractWorksets.IsEnabled = true;
                btnCancel.Content = "Cancel";
                pbProgress.IsIndeterminate = false;
                pbProgress.Value = 0;
            }
        }

        private async void BtnIntegrateTemplate_Click(object sender, RoutedEventArgs e)
        {
            if (_extractedFiles.Count == 0)
            {
                System.Windows.MessageBox.Show("No extracted files found for template integration.", "Error",
                    MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            // Ask user to select template file
            var openFileDialog = new Microsoft.Win32.OpenFileDialog
            {
                Filter = "Revit Files (*.rvt)|*.rvt|All Files (*.*)|*.*",
                Title = "Select Template File for Integration"
            };

            bool? result = openFileDialog.ShowDialog();
            if (result != true)
                return;

            string templateFilePath = openFileDialog.FileName;
            LogMessage($"Selected template file: {System.IO.Path.GetFileName(templateFilePath)}");

            // Disable UI during integration
            btnIntegrateTemplate.IsEnabled = false;
            btnQcCheck.IsEnabled = false;
            btnExtractWorksets.IsEnabled = false;
            btnCancel.Content = "Close";
            pbProgress.IsIndeterminate = true;
            txtStatus.Text = "Integrating into template...";

            try
            {
                LogMessage("=== TEMPLATE INTEGRATION STARTED ===");
                LogMessage($"Template file: {templateFilePath}");
                LogMessage($"Extracted files to integrate: {_extractedFiles.Count}");

                // Set parameters for template integration
                _eventHandler.SetTemplateIntegrationParameters(_orchestrator, _extractedFiles,
                    templateFilePath, _lastDestinationPath);

                // Raise the external event to execute in Revit context
                LogMessage("Starting template integration in Revit context...");
                _externalEvent.Raise();

                // Wait for completion
                await WaitForCompletionAsync();

                // Update UI based on result
                pbProgress.IsIndeterminate = false;
                pbProgress.Value = 100;

                if (_eventHandler.Success)
                {
                    LogMessage("=== TEMPLATE INTEGRATION COMPLETED SUCCESSFULLY ===");
                    txtStatus.Text = "Template integration completed";
                    System.Windows.MessageBox.Show($"Template integration completed successfully!\n" +
                        $"Integrated files are saved in the 'In Template' subfolder.", "Success",
                        MessageBoxButton.OK, MessageBoxImage.Information);
                }
                else
                {
                    LogMessage("=== TEMPLATE INTEGRATION FAILED ===");
                    txtStatus.Text = "Template integration failed";
                    if (_eventHandler.LastException != null)
                    {
                        LogMessage($"Error: {_eventHandler.LastException.Message}");
                    }
                    System.Windows.MessageBox.Show("Template integration failed. Check the log for details.",
                        "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
            catch (Exception ex)
            {
                LogMessage($"ERROR during template integration: {ex.Message}");
                txtStatus.Text = "Template integration error";
                System.Windows.MessageBox.Show($"An error occurred during template integration: {ex.Message}", "Error",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                btnIntegrateTemplate.IsEnabled = true;
                btnQcCheck.IsEnabled = true;
                btnExtractWorksets.IsEnabled = true;
                btnCancel.Content = "Cancel";
                pbProgress.IsIndeterminate = false;
                pbProgress.Value = 0;
            }
        }

        private List<string> GetExtractedFiles(string destinationPath)
        {
            var extractedFiles = new List<string>();

            try
            {
                if (Directory.Exists(destinationPath))
                {
                    // Look for RVT files that match the export pattern
                    var rvtFiles = Directory.GetFiles(destinationPath, "*_DX.rvt", SearchOption.TopDirectoryOnly);
                    extractedFiles.AddRange(rvtFiles);
                }
            }
            catch (Exception ex)
            {
                LogMessage($"Warning: Could not scan for extracted files: {ex.Message}");
            }

            return extractedFiles;
        }

        private async Task WaitForCompletionAsync()
        {
            // Wait for the external event to complete with a reasonable timeout
            int maxWaitTime = 600000; // 10 minutes for large models
            int checkInterval = 500; // 500ms
            int totalWaited = 0;
            int lastLogTime = 0;

            // Use determinate progress while waiting (progress = elapsed / max)
            pbProgress.IsIndeterminate = false;
            pbProgress.Minimum = 0;
            pbProgress.Maximum = maxWaitTime;
            pbProgress.Value = 0;
            txtStatus.Text = "Processing...";

            while (!_eventHandler.IsComplete && totalWaited < maxWaitTime)
            {
                await Task.Delay(checkInterval);
                totalWaited += checkInterval;

                // Update progress bar
                Dispatcher.Invoke(() =>
                {
                    if (pbProgress.Maximum > 0)
                        pbProgress.Value = Math.Min(totalWaited, (int)pbProgress.Maximum);
                });

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
                txtStatus.Text = "Timed out";
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