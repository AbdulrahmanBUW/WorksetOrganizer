using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace WorksetOrchestrator
{
    public class TemplateIntegrationService
    {
        private readonly UIDocument _uiDoc;
        private readonly Action<string> _logAction;
        private readonly WorksetMapper _worksetMapper;

        public TemplateIntegrationService(UIDocument uiDoc, Action<string> logAction, WorksetMapper worksetMapper)
        {
            _uiDoc = uiDoc;
            _logAction = logAction;
            _worksetMapper = worksetMapper;
        }

        public bool IntegrateIntoTemplate(List<string> extractedFiles, string templateFilePath, string destinationPath)
        {
            try
            {
                _logAction($"Starting template integration with {extractedFiles?.Count ?? 0} files...");
                _logAction($"Template file: {Path.GetFileName(templateFilePath)}");

                if (!File.Exists(templateFilePath))
                {
                    _logAction($"ERROR: Template file not found: {templateFilePath}");
                    return false;
                }

                string templateOutputPath = Path.Combine(destinationPath, "In Template");
                if (!Directory.Exists(templateOutputPath))
                {
                    Directory.CreateDirectory(templateOutputPath);
                    _logAction($"Created 'In Template' directory: {templateOutputPath}");
                }

                var app = _uiDoc.Application.Application;
                int processedCount = 0;

                foreach (var extractedFilePath in extractedFiles)
                {
                    _logAction($"=== Processing file {processedCount + 1}/{extractedFiles.Count}: {Path.GetFileName(extractedFilePath)} ===");

                    if (!File.Exists(extractedFilePath))
                    {
                        _logAction($"ERROR: Extracted file not found: {extractedFilePath}. Skipping.");
                        continue;
                    }

                    if (ProcessExtractedFile(extractedFilePath, templateFilePath, templateOutputPath, processedCount, app))
                    {
                        processedCount++;
                    }
                }

                _logAction($"=== TEMPLATE INTEGRATION COMPLETED ===");
                _logAction($"Successfully processed {processedCount} files.");
                _logAction($"Output location: {templateOutputPath}");

                return true;
            }
            catch (Exception ex)
            {
                _logAction($"ERROR in template integration: {ex.Message}");
                _logAction($"Stack Trace: {ex.StackTrace}");
                return false;
            }
        }

        private bool ProcessExtractedFile(string extractedFilePath, string templateFilePath, string templateOutputPath, int fileIndex, Application app)
        {
            string tempTemplatePath = null;
            Document templateDoc = null;
            Document extractedDoc = null;

            try
            {
                tempTemplatePath = Path.Combine(Path.GetDirectoryName(templateFilePath), $"temp_template_{fileIndex:000}_{Guid.NewGuid():N}.rvt");
                _logAction($"Creating temporary template copy: {tempTemplatePath}");
                File.Copy(templateFilePath, tempTemplatePath, true);
                File.SetAttributes(tempTemplatePath, FileAttributes.Normal);

                _logAction("Opening temporary template copy (attempting DetachAndPreserveWorksets)...");
                templateDoc = OpenTemplateDocument(tempTemplatePath, app);

                if (templateDoc == null)
                {
                    _logAction($"ERROR: Could not open temporary template copy: {tempTemplatePath}. Skipping file.");
                    CleanupTempFile(tempTemplatePath);
                    return false;
                }

                _logAction($"Opened template copy: {templateDoc.Title} (IsWorkshared: {templateDoc.IsWorkshared})");

                _logAction("Opening extracted file for reading...");
                try
                {
                    extractedDoc = app.OpenDocumentFile(extractedFilePath);
                }
                catch (Exception openExtractEx)
                {
                    _logAction($"ERROR opening extracted file: {openExtractEx.Message}");
                    extractedDoc = null;
                }

                if (extractedDoc == null)
                {
                    _logAction($"ERROR: Could not open extracted file: {extractedFilePath}. Closing template and skipping.");
                    try { templateDoc.Close(false); } catch { }
                    CleanupTempFile(tempTemplatePath);
                    return false;
                }

                _logAction($"Opened extracted file: {extractedDoc.Title}");

                var collectorService = new ElementCollectorService(extractedDoc, _logAction);
                var mepElements = collectorService.GetAllMepElementsFromDocument(extractedDoc);
                _logAction($"Found {mepElements.Count} MEP elements in extracted file.");

                if (mepElements.Count == 0)
                {
                    _logAction("No MEP elements found - closing docs and moving to next file.");
                    try { extractedDoc.Close(false); } catch { }
                    try { templateDoc.Close(false); } catch { }
                    CleanupTempFile(tempTemplatePath);
                    return false;
                }

                var elementIds = mepElements.Select(e => e.Id).ToList();

                ICollection<ElementId> copied = CopyElementsToTemplate(templateDoc, extractedDoc, elementIds, extractedFilePath);

                if (copied == null || !copied.Any())
                {
                    try { extractedDoc.Close(false); } catch { }
                    try { templateDoc.Close(false); } catch { }
                    CleanupTempFile(tempTemplatePath);
                    return false;
                }

                if (templateDoc.IsWorkshared)
                {
                    SynchronizeTemplateDocument(templateDoc);
                }

                string outputFileName = Path.GetFileName(extractedFilePath);
                string outputFilePath = Path.Combine(templateOutputPath, outputFileName);

                bool saved = SaveIntegratedDocument(templateDoc, outputFilePath);

                try { _logAction("Closing extracted file..."); extractedDoc.Close(false); } catch (Exception cex) { _logAction($"Warning closing extracted file: {cex.Message}"); }
                try { _logAction("Closing template copy..."); templateDoc.Close(false); } catch (Exception cex) { _logAction($"Warning closing template copy: {cex.Message}"); }

                CleanupTempFile(tempTemplatePath);

                return saved;
            }
            catch (Exception fileEx)
            {
                _logAction($"ERROR processing file '{Path.GetFileName(extractedFilePath)}': {fileEx.Message}");
                _logAction($"Stack trace: {fileEx.StackTrace}");
                try { if (extractedDoc != null) extractedDoc.Close(false); } catch { }
                try { if (templateDoc != null) templateDoc.Close(false); } catch { }
                CleanupTempFile(tempTemplatePath);
                return false;
            }
        }

        private Document OpenTemplateDocument(string tempTemplatePath, Application app)
        {
            try
            {
                OpenOptions openOpts = new OpenOptions();

                try
                {
                    openOpts.DetachFromCentralOption = DetachFromCentralOption.DetachAndPreserveWorksets;
                }
                catch
                {
                }

                ModelPath mp = null;
                try
                {
                    mp = ModelPathUtils.ConvertUserVisiblePathToModelPath(tempTemplatePath);
                }
                catch (Exception convEx)
                {
                    _logAction($"Warning: ModelPath conversion failed for '{tempTemplatePath}': {convEx.Message}");
                    mp = null;
                }

                if (mp != null)
                {
                    return app.OpenDocumentFile(mp, openOpts);
                }
                else
                {
                    return app.OpenDocumentFile(tempTemplatePath);
                }
            }
            catch (Exception openTempEx)
            {
                _logAction($"WARNING: Opening temp template failed ({openTempEx.Message}). Trying regular open...");
                try { return app.OpenDocumentFile(tempTemplatePath); }
                catch (Exception e2) { _logAction($"ERROR opening temp template: {e2.Message}"); return null; }
            }
        }

        private ICollection<ElementId> CopyElementsToTemplate(Document templateDoc, Document extractedDoc, List<ElementId> elementIds, string extractedFilePath)
        {
            _logAction("Starting copy transaction on template copy...");
            using (Transaction t = new Transaction(templateDoc, "Copy MEP elements into temp template"))
            {
                t.Start();
                try
                {
                    var copyOptions = new CopyPasteOptions();
                    copyOptions.SetDuplicateTypeNamesHandler(new DuplicateTypeNamesHandler());

                    var copied = ElementTransformUtils.CopyElements(
                        extractedDoc,
                        elementIds,
                        templateDoc,
                        Transform.Identity,
                        copyOptions);

                    _logAction($"Copied {copied?.Count ?? 0} elements into template copy.");

                    string sourceFileName = Path.GetFileNameWithoutExtension(extractedFilePath);
                    string worksetName = ExtractWorksetNameFromFileName(sourceFileName);

                    _logAction($"Extracted workset name '{worksetName}' from file name '{sourceFileName}'");

                    if (!string.IsNullOrEmpty(worksetName) && copied != null && copied.Any())
                    {
                        var worksetManager = new WorksetManager(templateDoc, _logAction);
                        worksetManager.AssignElementsToWorkset(templateDoc, copied.ToList(), worksetName);
                    }

                    t.Commit();
                    return copied;
                }
                catch (Exception copyEx)
                {
                    t.RollBack();
                    _logAction($"ERROR during copy: {copyEx.Message}");
                    return null;
                }
            }
        }

        private void SynchronizeTemplateDocument(Document templateDoc)
        {
            _logAction("Template copy is workshared - attempting SyncWithCentral + relinquish (safe-wrapped)...");
            try
            {
                var transOpts = new TransactWithCentralOptions();
                transOpts.SetLockCallback(new SynchLockCallback());

                var syncOpts = new SynchronizeWithCentralOptions();
                var relOpts = new RelinquishOptions(true);
                syncOpts.SetRelinquishOptions(relOpts);
                syncOpts.SaveLocalAfter = false;

                templateDoc.SynchronizeWithCentral(transOpts, syncOpts);
                _logAction("Synchronized and relinquished template copy.");
            }
            catch (Exception syncEx)
            {
                _logAction($"WARNING: SyncWithCentral failed on template copy: {syncEx.Message} — continuing to save.");
            }
        }

        private bool SaveIntegratedDocument(Document templateDoc, string outputFilePath)
        {
            _logAction($"Saving integrated file as: {outputFilePath}");
            try
            {
                var saveOpts = new SaveAsOptions { OverwriteExistingFile = true };

                if (templateDoc.IsWorkshared)
                {
                    try
                    {
                        var wsOpts = new WorksharingSaveAsOptions { SaveAsCentral = true };
                        saveOpts.SetWorksharingOptions(wsOpts);
                        _logAction("WorksharingSaveAsOptions.SaveAsCentral = true set for SaveAs (preserving worksets).");
                    }
                    catch (Exception exWs)
                    {
                        _logAction($"Warning: Could not set WorksharingSaveAsOptions: {exWs.Message}");
                    }
                }

                templateDoc.SaveAs(outputFilePath, saveOpts);
                _logAction($"Successfully saved integrated file: {outputFilePath}");
                return true;
            }
            catch (Exception saveEx)
            {
                _logAction($"ERROR saving integrated file: {saveEx.Message}");
                return false;
            }
        }

        private string ExtractWorksetNameFromFileName(string fileName)
        {
            try
            {
                var parts = fileName.Split('_');
                if (parts.Length >= 2)
                {
                    string iflsCode = parts[1];

                    var worksetToIflsMapping = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
                    {
                        {"DX_BDA", "B-D"},
                        {"DX_CDA", "D-D"},
                        {"DX_CHM", "C-S"},
                        {"DX_CKE", "C-L"},
                        {"DX_ELT", "E-X"},
                        {"DX_EXH", "A-X"},
                        {"DX_PAW", "S-D"},
                        {"DX_PG", "G-B"},
                        {"DX_PKW", "P-D"},
                        {"DX_PS", "G-S"},
                        {"DX_PWI", "U-D"},
                        {"DX_VAC", "V-D"},
                        {"DX_SLUR", "M-S"},
                        {"DX_STB", "S-T"},
                        {"DX_UPW", "U-D"},
                        {"DX_PVAC", "V-V"},
                        {"DX_RR", "R-R"},
                        {"DX_FND", "F-D"},
                        {"DX_Sub-tool", "S-BT"},
                        {"DX_Tool", "T-L"}
                    };

                    foreach (var kvp in worksetToIflsMapping)
                    {
                        if (kvp.Value.Equals(iflsCode, StringComparison.OrdinalIgnoreCase))
                        {
                            _logAction($"Mapped iFLS code '{iflsCode}' back to workset '{kvp.Key}'");
                            return kvp.Key;
                        }
                    }

                    _logAction($"Warning: Could not find workset for iFLS code '{iflsCode}'");
                }

                _logAction($"Warning: Could not extract iFLS code from filename '{fileName}'");
                return "DX_Unknown";
            }
            catch (Exception ex)
            {
                _logAction($"Error extracting workset name from filename '{fileName}': {ex.Message}");
                return "DX_Unknown";
            }
        }

        private void CleanupTempFile(string tempFilePath)
        {
            try
            {
                if (!string.IsNullOrEmpty(tempFilePath) && File.Exists(tempFilePath))
                {
                    File.SetAttributes(tempFilePath, FileAttributes.Normal);
                    File.Delete(tempFilePath);
                    _logAction($"Deleted temporary template copy: {tempFilePath}");
                }
            }
            catch (Exception exDelete) { _logAction($"Warning deleting temp template file: {exDelete.Message}"); }
        }

        private class SynchLockCallback : ICentralLockedCallback
        {
            public bool ShouldWaitForLockAvailability()
            {
                return false;
            }
        }
    }
}