using Autodesk.Revit.UI;
using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Interop;
using System.Windows.Forms; // Only for FolderBrowserDialog
using System.Collections.Generic;
using System.Windows.Media;
using System.Windows.Data;
using System.Globalization;

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
            InitializeInterface();
        }

        private void InitializeInterface()
        {
            txtStatus.Text = "Ready";
            pbProgress.Value = 0;
            txtProgress.Text = "0%";
            btnIntegrateTemplate.Visibility = Visibility.Collapsed;

            // Set placeholder text for TextBox (Excel file)
            SetPlaceholderText(txtExcelPath, "Select mapping Excel file...");

            // Set placeholder text for Label (Destination)
            SetPlaceholderLabel(lblDestination, "Select destination folder...");

            // Log initial information
            try
            {
                LogMessage($"Document: {_uiDoc.Document.Title}");
                LogMessage($"Workshared: {_uiDoc.Document.IsWorkshared}");
                LogMessage("Ready to organize worksets");
            }
            catch (Exception ex)
            {
                LogMessage($"Warning reading document info: {ex.Message}");
            }

            // Initialize in QC mode
            SetQcMode();
        }

        private void SetPlaceholderText(System.Windows.Controls.TextBox textBox, string placeholder)
        {
            textBox.Text = placeholder;
            textBox.Foreground = new SolidColorBrush(Color.FromRgb(102, 102, 102)); // #666666
            textBox.FontStyle = FontStyles.Italic;
            textBox.Tag = null; // No file selected initially
        }

        private void SetPlaceholderLabel(System.Windows.Controls.Label label, string placeholder)
        {
            label.Content = placeholder;
            label.Foreground = new SolidColorBrush(Color.FromRgb(102, 102, 102)); // #666666
            label.FontStyle = FontStyles.Italic;
            label.Tag = null; // No folder selected initially
        }

        private void SetSelectedText(System.Windows.Controls.TextBox textBox, string displayText, string fullPath)
        {
            textBox.Text = displayText;
            textBox.Foreground = new SolidColorBrush(Color.FromRgb(0, 0, 0)); // Black
            textBox.FontStyle = FontStyles.Normal;
            textBox.FontWeight = FontWeights.Bold;
            textBox.Tag = fullPath;
        }

        private void SetSelectedLabel(System.Windows.Controls.Label label, string displayText, string fullPath)
        {
            label.Content = displayText;
            label.Foreground = new SolidColorBrush(Color.FromRgb(0, 0, 0)); // Black
            label.FontStyle = FontStyles.Normal;
            label.FontWeight = FontWeights.Bold;
            label.Tag = fullPath;
        }

        protected override void OnClosed(EventArgs e)
        {
            _externalEvent?.Dispose();
            base.OnClosed(e);
        }

        private void SetQcMode()
        {
            _isExtractWorksetMode = false;

            // Update button styles
            btnQcMode.Style = (Style)FindResource("SegmentedButtonActive");
            btnExtractMode.Style = (Style)FindResource("SegmentedButton");

            // Update UI visibility
            panelExcelFile.Visibility = Visibility.Visible;
            panelOptions.Visibility = Visibility.Visible;

            // Update description
            txtModeDescription.Text = "Organizes and validates worksets based on Excel mapping. Elements are categorized by system patterns and moved to appropriate worksets before export.";

            // Update execute button
            btnExecute.Content = "Start QC & Extraction";

            LogMessage("Switched to QC Check & Extraction mode");
        }

        private void SetExtractWorksetsMode()
        {
            _isExtractWorksetMode = true;

            // Update button styles
            btnExtractMode.Style = (Style)FindResource("SegmentedButtonActive");
            btnQcMode.Style = (Style)FindResource("SegmentedButton");

            // Update UI visibility
            panelExcelFile.Visibility = Visibility.Collapsed;
            panelOptions.Visibility = Visibility.Collapsed;

            // Update description
            txtModeDescription.Text = "Extracts all available worksets into separate files. Preserves existing workset organization without reorganization.";

            // Update execute button
            btnExecute.Content = "Extract Worksets";

            LogMessage("Switched to Extract Worksets mode");
        }

        private void OnLogUpdated(object sender, string message)
        {
            Dispatcher.Invoke(() =>
            {
                txtLog.AppendText(message + Environment.NewLine);
                txtLog.ScrollToEnd();

                // Update status
                if (!string.IsNullOrEmpty(message))
                {
                    int hyphenIndex = message.IndexOf(" - ");
                    string shortMsg = message;
                    if (hyphenIndex >= 0 && hyphenIndex + 3 < message.Length)
                        shortMsg = message.Substring(hyphenIndex + 3);

                    txtStatus.Text = shortMsg.Length > 60 ? shortMsg.Substring(0, 57) + "..." : shortMsg;
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
                // Update the TextBox for Excel file selection
                SetSelectedText(txtExcelPath, Path.GetFileName(openFileDialog.FileName), openFileDialog.FileName);
                LogMessage($"Selected Excel file: {Path.GetFileName(openFileDialog.FileName)}");
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
                // Update the Label for destination folder selection
                SetSelectedLabel(lblDestination, Path.GetFileName(folderDialog.SelectedPath), folderDialog.SelectedPath);
                LogMessage($"Selected destination: {folderDialog.SelectedPath}");
            }
        }

        private void BtnQcCheck_Click(object sender, RoutedEventArgs e)
        {
            if (!_isExtractWorksetMode) return; // Already in QC mode
            SetQcMode();
        }

        private void BtnExtractWorksets_Click(object sender, RoutedEventArgs e)
        {
            if (_isExtractWorksetMode) return; // Already in Extract mode
            SetExtractWorksetsMode();
        }

        private async void BtnExecute_Click(object sender, RoutedEventArgs e)
        {
            if (_isExtractWorksetMode)
            {
                await RunExtractWorksets();
            }
            else
            {
                await RunQcCheckAndExtraction();
            }
        }

        private async Task RunQcCheckAndExtraction()
        {
            // Validation for QC mode - get from TextBox
            string excelPath = txtExcelPath.Tag as string;
            if (string.IsNullOrEmpty(excelPath) || !File.Exists(excelPath))
            {
                ShowAlert("Please select a valid Excel file.", "Configuration Required");
                return;
            }

            // Validation for destination - get from Label
            string destinationPath = lblDestination.Tag as string;
            if (string.IsNullOrEmpty(destinationPath) || !Directory.Exists(destinationPath))
            {
                ShowAlert("Please select a valid destination folder.", "Configuration Required");
                return;
            }

            if (!_uiDoc.Document.IsWorkshared)
            {
                ShowAlert("The current document is not workshared. Please enable worksharing first.", "Worksharing Required");
                return;
            }

            await ExecuteOperation("QC Check & Extraction", async () =>
            {
                LogMessage("=== QC CHECK & EXTRACTION STARTED ===");
                LogMessage("Reading Excel mapping file...");

                var mapping = ExcelReader.ReadMapping(excelPath);
                LogMessage($"Found {mapping.Count} mapping records");

                // Log mapping summary
                foreach (var record in mapping.Take(3))
                {
                    LogMessage($"  Pattern: '{record.SystemNameInModel}' → Workset: '{record.WorksetName}' → Package: '{record.ModelPackageCode}'");
                }
                if (mapping.Count > 3)
                    LogMessage($"  ... and {mapping.Count - 3} more records");

                _eventHandler.SetParameters(_orchestrator, mapping, destinationPath,
                    chkOverwrite.IsChecked == true, chkExportQc.IsChecked == true);

                LogMessage("Starting workset orchestration...");
                _externalEvent.Raise();

                await WaitForCompletionAsync();
            });
        }

        private async Task RunExtractWorksets()
        {
            // Validation for destination - get from Label
            string destinationPath = lblDestination.Tag as string;
            if (string.IsNullOrEmpty(destinationPath) || !Directory.Exists(destinationPath))
            {
                ShowAlert("Please select a valid destination folder.", "Configuration Required");
                return;
            }

            if (!_uiDoc.Document.IsWorkshared)
            {
                ShowAlert("The current document is not workshared. Please enable worksharing first.", "Worksharing Required");
                return;
            }

            await ExecuteOperation("Extract Worksets", async () =>
            {
                LogMessage("=== EXTRACT WORKSETS STARTED ===");
                LogMessage("Extracting all available worksets...");

                _eventHandler.SetWorksetExtractionParameters(_orchestrator, destinationPath, true);

                LogMessage("Starting workset extraction...");
                _externalEvent.Raise();

                await WaitForCompletionAsync();
            });
        }

        private async Task ExecuteOperation(string operationName, Func<Task> operation)
        {
            _extractedFiles.Clear();
            btnIntegrateTemplate.Visibility = Visibility.Collapsed;
            _lastDestinationPath = lblDestination.Tag as string; // Updated to use Label

            // Disable UI
            btnExecute.IsEnabled = false;
            btnQcMode.IsEnabled = false;
            btnExtractMode.IsEnabled = false;
            btnCancel.Content = "Close";
            txtLog.Clear();

            // Show progress
            UpdateProgress(0, "Initializing...");

            try
            {
                await operation();

                if (_eventHandler.Success)
                {
                    UpdateProgress(100, "Completed successfully");
                    LogMessage($"=== {operationName.ToUpper()} COMPLETED SUCCESSFULLY ===");

                    _extractedFiles = GetExtractedFiles(_lastDestinationPath);
                    LogMessage($"Found {_extractedFiles.Count} extracted files");

                    if (_extractedFiles.Count > 0)
                    {
                        btnIntegrateTemplate.Visibility = Visibility.Visible;
                        LogMessage("Template integration is now available");
                    }

                    ShowSuccess($"{operationName} completed successfully! Generated {_extractedFiles.Count} files.");
                }
                else
                {
                    UpdateProgress(0, "Completed with errors");
                    LogMessage($"=== {operationName.ToUpper()} COMPLETED WITH ERRORS ===");

                    if (_eventHandler.LastException != null)
                    {
                        LogMessage($"Error: {_eventHandler.LastException.Message}");
                    }

                    ShowAlert($"{operationName} completed with errors. Check the log for details.", "Processing Completed");
                }
            }
            catch (Exception ex)
            {
                UpdateProgress(0, "Error occurred");
                LogMessage($"ERROR: {ex.Message}");
                ShowAlert($"An error occurred: {ex.Message}", "Error");
            }
            finally
            {
                // Re-enable UI
                btnExecute.IsEnabled = true;
                btnQcMode.IsEnabled = true;
                btnExtractMode.IsEnabled = true;
                btnCancel.Content = "Cancel";

                if (pbProgress.Value == 0)
                {
                    UpdateProgress(0, "Ready");
                }
            }
        }

        private async void BtnIntegrateTemplate_Click(object sender, RoutedEventArgs e)
        {
            if (_extractedFiles.Count == 0)
            {
                ShowAlert("No extracted files found for template integration.", "No Files Available");
                return;
            }

            var openFileDialog = new Microsoft.Win32.OpenFileDialog
            {
                Filter = "Revit Files (*.rvt)|*.rvt|All Files (*.*)|*.*",
                Title = "Select Template File for Integration"
            };

            bool? result = openFileDialog.ShowDialog();
            if (result != true) return;

            string templateFilePath = openFileDialog.FileName;
            LogMessage($"Selected template: {Path.GetFileName(templateFilePath)}");

            // Disable UI during integration
            btnIntegrateTemplate.IsEnabled = false;
            btnExecute.IsEnabled = false;
            btnQcMode.IsEnabled = false;
            btnExtractMode.IsEnabled = false;
            btnCancel.Content = "Close";

            UpdateProgress(0, "Integrating into template...");

            try
            {
                LogMessage("=== TEMPLATE INTEGRATION STARTED ===");
                LogMessage($"Template: {templateFilePath}");
                LogMessage($"Files to integrate: {_extractedFiles.Count}");

                _eventHandler.SetTemplateIntegrationParameters(_orchestrator, _extractedFiles,
                    templateFilePath, _lastDestinationPath);

                LogMessage("Starting template integration...");
                _externalEvent.Raise();

                await WaitForCompletionAsync();

                if (_eventHandler.Success)
                {
                    UpdateProgress(100, "Template integration completed");
                    LogMessage("=== TEMPLATE INTEGRATION COMPLETED SUCCESSFULLY ===");
                    ShowSuccess("Template integration completed successfully!\nIntegrated files are saved in the 'In Template' subfolder.");
                }
                else
                {
                    UpdateProgress(0, "Template integration failed");
                    LogMessage("=== TEMPLATE INTEGRATION FAILED ===");
                    if (_eventHandler.LastException != null)
                    {
                        LogMessage($"Error: {_eventHandler.LastException.Message}");
                    }
                    ShowAlert("Template integration failed. Check the log for details.", "Integration Failed");
                }
            }
            catch (Exception ex)
            {
                UpdateProgress(0, "Template integration error");
                LogMessage($"ERROR during template integration: {ex.Message}");
                ShowAlert($"An error occurred during template integration: {ex.Message}", "Error");
            }
            finally
            {
                btnIntegrateTemplate.IsEnabled = true;
                btnExecute.IsEnabled = true;
                btnQcMode.IsEnabled = true;
                btnExtractMode.IsEnabled = true;
                btnCancel.Content = "Cancel";
            }
        }

        private List<string> GetExtractedFiles(string destinationPath)
        {
            var extractedFiles = new List<string>();

            try
            {
                if (Directory.Exists(destinationPath))
                {
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
            int maxWaitTime = 600000; // 10 minutes
            int checkInterval = 500;
            int totalWaited = 0;
            int lastLogTime = 0;

            while (!_eventHandler.IsComplete && totalWaited < maxWaitTime)
            {
                await Task.Delay(checkInterval);
                totalWaited += checkInterval;

                // Update progress
                Dispatcher.Invoke(() =>
                {
                    double progressPercent = Math.Min((double)totalWaited / maxWaitTime * 100, 90);
                    UpdateProgress(progressPercent, "Processing...");
                });

                // Log progress every 30 seconds
                if (totalWaited - lastLogTime >= 30000)
                {
                    LogMessage($"Processing... ({totalWaited / 1000} seconds elapsed)");
                    lastLogTime = totalWaited;
                }

                System.Windows.Forms.Application.DoEvents();
            }

            if (totalWaited >= maxWaitTime)
            {
                LogMessage("WARNING: Operation timed out after 10 minutes");
                UpdateProgress(0, "Timed out");
            }
        }

        private void UpdateProgress(double value, string status)
        {
            pbProgress.Value = value;
            txtProgress.Text = $"{value:F0}%";
            txtStatus.Text = status;
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

        private void ShowAlert(string message, string title)
        {
            System.Windows.MessageBox.Show(message, title, MessageBoxButton.OK, MessageBoxImage.Warning);
        }

        private void ShowSuccess(string message)
        {
            System.Windows.MessageBox.Show(message, "Success", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        private void txtExcelPath_TextChanged(object sender, System.Windows.Controls.TextChangedEventArgs e)
        {
            // Empty event handler - required by XAML
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