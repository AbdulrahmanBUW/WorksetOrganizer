using Autodesk.Revit.UI;
using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Interop;
using System.Windows.Forms;
using System.Collections.Generic;
using System.Windows.Media;
using System.Windows.Controls;

namespace WorksetOrchestrator
{
    public partial class MainForm : Window
    {
        private readonly UIDocument _uiDoc;
        private readonly WorksetOrchestrator _orchestrator;
        private readonly WorksetEventHandler _eventHandler;
        private readonly ExternalEvent _externalEvent;
        private readonly List<string> _extractedFiles = new List<string>();
        private string _lastDestinationPath;
        private bool _isExtractWorksetMode = false;
        private readonly List<System.Windows.Controls.CheckBox> _worksetCheckboxes = new List<System.Windows.Controls.CheckBox>();

        public MainForm(UIDocument uiDoc, IntPtr revitMainWindowHandle)
        {
            InitializeComponent();

            _uiDoc = uiDoc;
            _orchestrator = new WorksetOrchestrator(uiDoc);
            _orchestrator.LogUpdated += OnLogUpdated;

            _eventHandler = new WorksetEventHandler();
            _externalEvent = ExternalEvent.Create(_eventHandler);

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
            }

            InitializeInterface();
        }

        private void InitializeInterface()
        {
            txtStatus.Text = "Ready";
            pbProgress.Value = 0;
            txtProgress.Text = "0%";
            btnIntegrateTemplate.Visibility = Visibility.Collapsed;

            SetPlaceholderText(txtExcelPath, "Select mapping Excel file...");

            SetPlaceholderLabel(lblDestination, "Select destination folder...");

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

            SetQcMode();
        }

        private void SetPlaceholderText(System.Windows.Controls.TextBox textBox, string placeholder)
        {
            textBox.Text = placeholder;
            textBox.Foreground = new SolidColorBrush(Color.FromRgb(102, 102, 102));
            textBox.FontStyle = FontStyles.Italic;
            textBox.Tag = null;
        }

        private void SetPlaceholderLabel(System.Windows.Controls.Label label, string placeholder)
        {
            label.Content = placeholder;
            label.Foreground = new SolidColorBrush(Color.FromRgb(102, 102, 102));
            label.FontStyle = FontStyles.Italic;
            label.Tag = null;
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

            btnQcMode.Style = (Style)FindResource("SegmentedButtonActive");
            btnExtractMode.Style = (Style)FindResource("SegmentedButton");

            panelExcelFile.Visibility = Visibility.Visible;
            panelOptions.Visibility = Visibility.Visible;
            panelWorksetSelection.Visibility = Visibility.Collapsed;

            txtModeDescription.Text = "Organizes and validates worksets based on Excel mapping. Elements are categorized by system patterns and moved to appropriate worksets before export.";

            btnExecute.Content = "Start QC & Extraction";

            LogMessage("Switched to QC Check & Extraction mode");
        }

        private void PopulateWorksetCheckboxes()
        {
            worksetCheckboxContainer.Children.Clear();
            _worksetCheckboxes.Clear();

            try
            {
                var availableWorksets = _orchestrator.GetAvailableWorksets();

                if (availableWorksets.Count == 0)
                {
                    var noWorksetsText = new TextBlock
                    {
                        Text = "No worksets available in this document",
                        FontSize = 14,
                        FontStyle = FontStyles.Italic,
                        Foreground = new SolidColorBrush(Color.FromRgb(142, 142, 147)),
                        Margin = new Thickness(0, 8, 0, 8)
                    };
                    worksetCheckboxContainer.Children.Add(noWorksetsText);
                    LogMessage("No worksets available for selection");
                    return;
                }

                LogMessage($"Found {availableWorksets.Count} worksets available for selection");

                foreach (var worksetName in availableWorksets.OrderBy(w => w))
                {
                    // CORRECTED: Fully qualify the CheckBox type
                    var checkbox = new System.Windows.Controls.CheckBox
                    {
                        Content = worksetName,
                        IsChecked = true, // Default to all selected
                        Margin = new Thickness(0, 4, 0, 4),
                        FontSize = 14,
                        FontFamily = new System.Windows.Media.FontFamily("Segoe UI")
                    };

                    _worksetCheckboxes.Add(checkbox);
                    worksetCheckboxContainer.Children.Add(checkbox);
                }

                LogMessage($"Loaded {_worksetCheckboxes.Count} worksets (all selected by default)");
            }
            catch (Exception ex)
            {
                LogMessage($"Error loading worksets: {ex.Message}");
                ShowAlert("Could not load worksets from the document.", "Error");
            }
        }

        // NEW: Select/Deselect All buttons
        private void BtnSelectAllWorksets_Click(object sender, RoutedEventArgs e)
        {
            foreach (var checkbox in _worksetCheckboxes)
            {
                checkbox.IsChecked = true;
            }
            LogMessage("All worksets selected");
        }
        private void BtnDeselectAllWorksets_Click(object sender, RoutedEventArgs e)
        {
            foreach (var checkbox in _worksetCheckboxes)
            {
                checkbox.IsChecked = false;
            }
            LogMessage("All worksets deselected");
        }

        // NEW: Get selected worksets
        private List<string> GetSelectedWorksets()
        {
            return _worksetCheckboxes
                .Where(cb => cb.IsChecked == true)
                .Select(cb => cb.Content.ToString())
                .ToList();
        }


        private void SetExtractWorksetsMode()
        {
            _isExtractWorksetMode = true;

            btnExtractMode.Style = (Style)FindResource("SegmentedButtonActive");
            btnQcMode.Style = (Style)FindResource("SegmentedButton");

            panelExcelFile.Visibility = Visibility.Collapsed;
            panelOptions.Visibility = Visibility.Collapsed;
            panelWorksetSelection.Visibility = Visibility.Visible;

            txtModeDescription.Text = "Extracts selected worksets into separate files. Choose which worksets to extract using the checkboxes below.";

            btnExecute.Content = "Extract Worksets";

            PopulateWorksetCheckboxes();

            LogMessage("Switched to Extract Worksets mode");
        }

        private void OnLogUpdated(object sender, string message)
        {
            Dispatcher.Invoke(() =>
            {
                txtLog.AppendText(message + Environment.NewLine);
                txtLog.ScrollToEnd();

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
                SetSelectedLabel(lblDestination, Path.GetFileName(folderDialog.SelectedPath), folderDialog.SelectedPath);
                LogMessage($"Selected destination: {folderDialog.SelectedPath}");
            }
        }

        private void BtnQcCheck_Click(object sender, RoutedEventArgs e)
        {
            if (!_isExtractWorksetMode) return;
            SetQcMode();
        }

        private void BtnExtractWorksets_Click(object sender, RoutedEventArgs e)
        {
            if (_isExtractWorksetMode) return;
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
            string excelPath = txtExcelPath.Tag as string;
            if (string.IsNullOrEmpty(excelPath) || !File.Exists(excelPath))
            {
                ShowAlert("Please select a valid Excel file.", "Configuration Required");
                return;
            }

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

            var selectedWorksets = GetSelectedWorksets();

            if (selectedWorksets.Count == 0)
            {
                ShowAlert("Please select at least one workset to extract.", "No Worksets Selected");
                return;
            }

            await ExecuteOperation("Extract Worksets", async () =>
            {
                LogMessage("=== EXTRACT WORKSETS STARTED ===");
                LogMessage($"Extracting {selectedWorksets.Count} selected worksets...");
                LogMessage($"Selected: {string.Join(", ", selectedWorksets)}");

                _eventHandler.SetWorksetExtractionParameters(_orchestrator, destinationPath, true, selectedWorksets);

                LogMessage("Starting workset extraction...");
                _externalEvent.Raise();

                await WaitForCompletionAsync();
            });
        }

        private async Task ExecuteOperation(string operationName, Func<Task> operation)
        {
            _extractedFiles.Clear();
            btnIntegrateTemplate.Visibility = Visibility.Collapsed;
            _lastDestinationPath = lblDestination.Tag as string;

            btnExecute.IsEnabled = false;
            btnQcMode.IsEnabled = false;
            btnExtractMode.IsEnabled = false;
            btnCancel.Content = "Close";
            txtLog.Clear();

            UpdateProgress(0, "Initializing...");

            try
            {
                await operation();

                if (_eventHandler.Success)
                {
                    UpdateProgress(100, "Completed successfully");
                    LogMessage($"=== {operationName.ToUpper()} COMPLETED SUCCESSFULLY ===");

                    _extractedFiles.Clear();
                    _extractedFiles.AddRange(GetExtractedFiles(_lastDestinationPath));
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
            int maxWaitTime = 600000;
            int checkInterval = 500;
            int totalWaited = 0;
            int lastLogTime = 0;

            while (!_eventHandler.IsComplete && totalWaited < maxWaitTime)
            {
                await Task.Delay(checkInterval);
                totalWaited += checkInterval;

                Dispatcher.Invoke(() =>
                {
                    double progressPercent = Math.Min((double)totalWaited / maxWaitTime * 100, 90);
                    UpdateProgress(progressPercent, "Processing...");
                });

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

        private void TxtExcelPath_TextChanged(object sender, System.Windows.Controls.TextChangedEventArgs e)
        {
        }
    }

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