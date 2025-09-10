using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.ApplicationServices;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.IO;

namespace WorksetOrchestrator
{
    public class WorksetOrchestrator
    {
        private Document _doc;
        private UIDocument _uiDoc;
        private StringBuilder _log = new StringBuilder();

        public event EventHandler<string> LogUpdated;

        public WorksetOrchestrator(UIDocument uiDoc)
        {
            _uiDoc = uiDoc;
            _doc = uiDoc.Document;
        }

        public string Log => _log.ToString();

        public void LogMessage(string message)
        {
            string formattedMessage = $"{DateTime.Now:HH:mm:ss} - {message}";
            _log.AppendLine(formattedMessage);
            LogUpdated?.Invoke(this, formattedMessage);
        }

        public bool Execute(List<MappingRecord> mapping, string destinationPath, bool overwriteFiles, bool exportQc)
        {
            try
            {
                if (!_doc.IsWorkshared)
                {
                    LogMessage("ERROR: Document is not workshared. Please enable worksharing.");
                    return false;
                }

                var packageGroupMapping = new Dictionary<string, List<ElementId>>();
                
                var eltRecord = mapping.FirstOrDefault(m => m.WorksetName?.ToUpper() == "DX_ELT");
                if (eltRecord == null)
                {
                    // Add a default ELT record if missing
                    mapping.Add(new MappingRecord
                    {
                        WorksetName = "DX_ELT",
                        SystemNameInModel = "ELT",
                        SystemDescription = "Electrical Elements",
                        ModelPackageCode = "ELT" // or whatever your code should be
                    });
                    LogMessage("Added default DX_ELT mapping record");
                }

                using (Transaction trans = new Transaction(_doc, "Process Workset Mapping"))
                {
                    trans.Start();

                    // Create DX_QC workset inside transaction
                    WorksetId qcWorksetId = GetOrCreateWorkset("DX_QC");

                    foreach (var record in mapping)
                    {
                        LogMessage($"Processing record for system pattern: {record.SystemNameInModel}");

                        if (record.ModelPackageCode == "NO EXPORT")
                        {
                            LogMessage($"Skipped - Marked as 'NO EXPORT'.");
                            continue;
                        }

                        // Create/get target workset inside transaction
                        WorksetId targetWorksetId = GetOrCreateWorkset(record.WorksetName);

                        // Get all MEP elements that could have system information
                        var allMepElements = GetAllMepElements();

                        // Find matching elements using improved pattern matching
                        var matchingElements = FindMatchingElements(allMepElements, record);

                        LogMessage($"Found {matchingElements.Count} elements for system pattern '{record.SystemNameInModel}'.");

                        int movedCount = 0;
                        foreach (Element element in matchingElements)
                        {
                            var worksetParam = element.get_Parameter(BuiltInParameter.ELEM_PARTITION_PARAM);
                            if (worksetParam != null && !worksetParam.IsReadOnly && element.WorksetId != targetWorksetId)
                            {
                                worksetParam.Set(targetWorksetId.IntegerValue);
                                movedCount++;
                            }
                        }

                        LogMessage($"Moved {movedCount} elements to workset '{record.WorksetName}'.");

                        // Only add to package mapping if we found elements
                        if (matchingElements.Count > 0)
                        {
                            string normalizedCode = record.NormalizedPackageCode;
                            if (!packageGroupMapping.ContainsKey(normalizedCode))
                                packageGroupMapping[normalizedCode] = new List<ElementId>();

                            packageGroupMapping[normalizedCode].AddRange(matchingElements.Select(e => e.Id));
                        }
                    }



                    // Move remaining MEP elements to DX_QC
                    var allMepElementsAfterProcessing = GetAllMepElements();
                    var processedElementIds = packageGroupMapping.Values.SelectMany(x => x).ToHashSet();


                    var cableTrays = allMepElementsAfterProcessing.Where(e =>
                    e.Category?.Id.IntegerValue == (int)BuiltInCategory.OST_CableTray ||
                    e.Category?.Id.IntegerValue == (int)BuiltInCategory.OST_CableTrayFitting).ToList();

                    LogMessage($"Found {cableTrays.Count} cable trays/fittings after processing");
                    foreach (var ct in cableTrays.Take(5)) // Log first 5
                    {
                        bool isProcessed = processedElementIds.Contains(ct.Id);
                        LogMessage($"  Cable Tray {ct.Id}: Processed = {isProcessed}");
                    }

                    int orphanedCount = 0;
                    foreach (var element in allMepElementsAfterProcessing)
                    {
                        if (!processedElementIds.Contains(element.Id))
                        {
                            var worksetParam = element.get_Parameter(BuiltInParameter.ELEM_PARTITION_PARAM);
                            if (worksetParam != null && !worksetParam.IsReadOnly && element.WorksetId != qcWorksetId)
                            {
                                worksetParam.Set(qcWorksetId.IntegerValue);
                                orphanedCount++;
                            }
                        }
                    }

                    LogMessage($"Moved {orphanedCount} orphaned MEP elements to 'DX_QC' workset.");

                    trans.Commit();
                }

                // Export logic outside transaction - only export if elements exist
                ExportRVTs(packageGroupMapping, destinationPath, overwriteFiles, exportQc);

                // Write final log
                System.IO.File.WriteAllText(System.IO.Path.Combine(destinationPath, "WorksetOrchestrationLog.txt"), _log.ToString());

                return true;
            }
            catch (Exception ex)
            {
                LogMessage($"ERROR: {ex.Message}");
                LogMessage($"Stack Trace: {ex.StackTrace}");
                return false;
            }
        }

        /// <summary>
        /// New method: Extract all available worksets into separate files
        /// No QC checking, no Excel mapping - just extract what exists
        /// </summary>
        public bool ExecuteWorksetExtraction(string destinationPath, bool overwriteFiles)
        {
            try
            {
                if (!_doc.IsWorkshared)
                {
                    LogMessage("ERROR: Document is not workshared. Please enable worksharing.");
                    return false;
                }

                LogMessage("=== WORKSET EXTRACTION MODE ===");
                LogMessage("Analyzing available worksets...");

                // Get all user worksets
                var allWorksets = new FilteredWorksetCollector(_doc)
                    .OfKind(WorksetKind.UserWorkset)
                    .Where(w => w.Kind == WorksetKind.UserWorkset)
                    .ToList();

                LogMessage($"Found {allWorksets.Count} user worksets in the model:");
                foreach (var workset in allWorksets)
                {
                    LogMessage($"  - {workset.Name} (ID: {workset.Id})");
                }

                // Create workset-to-elements mapping
                var worksetElementMapping = new Dictionary<string, List<ElementId>>();

                foreach (var workset in allWorksets)
                {
                    LogMessage($"Collecting elements for workset: {workset.Name}");

                    // Get all elements in this workset
                    var elementsInWorkset = new FilteredElementCollector(_doc)
                        .WhereElementIsNotElementType()
                        .Where(e => e.WorksetId == workset.Id)
                        .ToList();

                    // Filter to get only relevant elements (MEP, structural, etc.)
                    var relevantElements = GetRelevantElementsFromCollection(elementsInWorkset);

                    LogMessage($"  Found {elementsInWorkset.Count} total elements, {relevantElements.Count} relevant elements");

                    if (relevantElements.Count > 0)
                    {
                        worksetElementMapping[workset.Name] = relevantElements.Select(e => e.Id).ToList();

                        // Log sample elements for verification
                        var sampleElements = relevantElements.Take(3).ToList();
                        foreach (var elem in sampleElements)
                        {
                            string categoryName = elem.Category?.Name ?? "Unknown";
                            LogMessage($"    Sample: {elem.Id} ({categoryName})");
                        }
                    }
                    else
                    {
                        LogMessage($"    No relevant elements found in workset '{workset.Name}' - skipping");
                    }
                }

                LogMessage($"Worksets with relevant elements: {worksetElementMapping.Count}");

                // Synchronize with central before export (same as QC mode)
                if (_doc.IsWorkshared)
                {
                    try
                    {
                        LogMessage("Synchronizing with central and relinquishing ownership...");

                        var transOpts = new TransactWithCentralOptions();
                        transOpts.SetLockCallback(new SynchLockCallback());

                        var syncOpts = new SynchronizeWithCentralOptions();
                        var relinquishOpts = new RelinquishOptions(true);
                        syncOpts.SetRelinquishOptions(relinquishOpts);
                        syncOpts.SaveLocalAfter = false;

                        _doc.SynchronizeWithCentral(transOpts, syncOpts);
                        LogMessage("Synchronized with central and relinquished ownership.");
                    }
                    catch (Exception ex)
                    {
                        LogMessage($"WARNING: Synchronization with central failed: {ex.Message}");
                    }
                }

                // Export worksets as separate RVT files
                ExportWorksetRVTs(worksetElementMapping, destinationPath, overwriteFiles);

                // Write final log
                System.IO.File.WriteAllText(System.IO.Path.Combine(destinationPath, "WorksetExtractionLog.txt"), _log.ToString());

                LogMessage("=== WORKSET EXTRACTION COMPLETED ===");
                return true;
            }
            catch (Exception ex)
            {
                LogMessage($"ERROR: {ex.Message}");
                LogMessage($"Stack Trace: {ex.StackTrace}");
                return false;
            }
        }

        /// <summary>
        /// Filter elements to include only relevant categories for extraction
        /// </summary>
        private List<Element> GetRelevantElementsFromCollection(List<Element> elements)
        {
            var relevantElements = new List<Element>();

            // Define categories we want to include in extraction
            var relevantCategories = new List<BuiltInCategory>
    {
        // MEP Categories
        BuiltInCategory.OST_PipeFitting,
        BuiltInCategory.OST_PipeAccessory,
        BuiltInCategory.OST_PipeCurves,
        BuiltInCategory.OST_DuctFitting,
        BuiltInCategory.OST_DuctAccessory,
        BuiltInCategory.OST_DuctCurves,
        BuiltInCategory.OST_DuctTerminal,
        BuiltInCategory.OST_MechanicalEquipment,
        BuiltInCategory.OST_PlumbingFixtures,
        BuiltInCategory.OST_Sprinklers,
        BuiltInCategory.OST_ElectricalEquipment,
        BuiltInCategory.OST_ElectricalFixtures,
        BuiltInCategory.OST_LightingFixtures,
        BuiltInCategory.OST_CableTray,
        BuiltInCategory.OST_CableTrayFitting,
        BuiltInCategory.OST_Conduit,
        BuiltInCategory.OST_ConduitFitting,
        
        // Structural Categories
        BuiltInCategory.OST_StructuralFraming,
        BuiltInCategory.OST_StructuralColumns,
        BuiltInCategory.OST_StructuralFoundation,
        BuiltInCategory.OST_StructuralFramingSystem,
        BuiltInCategory.OST_StructuralStiffener,
        BuiltInCategory.OST_StructuralTruss,
        
        // Architectural Categories (selective)
        BuiltInCategory.OST_Walls,
        BuiltInCategory.OST_Doors,
        BuiltInCategory.OST_Windows,
        
        // Generic and Specialty
        BuiltInCategory.OST_GenericModel,
        BuiltInCategory.OST_SpecialityEquipment
    };

            foreach (var element in elements)
            {
                try
                {
                    if (element.Category != null)
                    {
                        var categoryId = (BuiltInCategory)element.Category.Id.IntegerValue;
                        if (relevantCategories.Contains(categoryId))
                        {
                            relevantElements.Add(element);
                        }
                    }
                }
                catch (Exception ex)
                {
                    LogMessage($"Warning: Error checking element {element.Id}: {ex.Message}");
                }
            }

            return relevantElements;
        }

        /// <summary>
        /// Export worksets as individual RVT files
        /// Similar to ExportRVTs but organized by workset instead of by mapping
        /// </summary>
        private void ExportWorksetRVTs(Dictionary<string, List<ElementId>> worksetMapping, string destinationPath, bool overwrite)
        {
            string projectPrefix = System.IO.Path.GetFileNameWithoutExtension(_doc.PathName);
            if (projectPrefix.Contains('_'))
                projectPrefix = projectPrefix.Split('_')[0];

            LogMessage($"Preparing to export {worksetMapping.Count} workset files...");

            int exportedCount = 0;
            var app = _uiDoc.Application.Application;

            foreach (var worksetGroup in worksetMapping)
            {
                string worksetName = worksetGroup.Key;
                var elementIds = worksetGroup.Value;

                if (elementIds == null || !elementIds.Any())
                {
                    LogMessage($"Skipping '{worksetName}' - no elements to export");
                    continue;
                }

                // Clean workset name for file system
                string cleanWorksetName = CleanFileSystemName(worksetName);
                string exportFileName = $"{projectPrefix}_{cleanWorksetName}_WS_DX.rvt";
                string exportFilePath = System.IO.Path.Combine(destinationPath, exportFileName);

                // If file exists and we should not overwrite, skip.
                if (System.IO.File.Exists(exportFilePath) && !overwrite)
                {
                    LogMessage($"Skipped - File exists: {exportFileName}");
                    continue;
                }

                Document newDoc = null;
                try
                {
                    // Create new project document
                    LogMessage($"Creating new document for workset: {worksetName}");

                    string templatePath = null;
                    try
                    {
                        templatePath = _uiDoc.Application.Application.DefaultProjectTemplate;
                    }
                    catch { templatePath = null; }

                    if (string.IsNullOrEmpty(templatePath) || !File.Exists(templatePath))
                    {
                        newDoc = app.NewProjectDocument(UnitSystem.Metric);
                    }
                    else
                    {
                        newDoc = app.NewProjectDocument(templatePath);
                    }

                    if (newDoc == null)
                        throw new Exception("Failed to create new project document.");

                    LogMessage($"Created temporary document for workset '{worksetName}' with {elementIds.Count} elements");

                    // Copy elements to new document
                    try
                    {
                        using (Transaction tNew = new Transaction(newDoc, $"Copy workset {worksetName}"))
                        {
                            tNew.Start();

                            var copyOptions = new CopyPasteOptions();
                            copyOptions.SetDuplicateTypeNamesHandler(new DuplicateTypeNamesHandler());

                            ICollection<ElementId> copiedIds = ElementTransformUtils.CopyElements(
                                _doc,
                                elementIds,
                                newDoc,
                                Transform.Identity,
                                copyOptions
                            );

                            tNew.Commit();

                            LogMessage($"Copied {copiedIds?.Count ?? 0} elements for workset '{worksetName}'");
                        }
                    }
                    catch (Exception copyEx)
                    {
                        LogMessage($"ERROR copying elements for workset '{worksetName}': {copyEx.Message}");
                        try { newDoc.Close(false); } catch { }
                        continue;
                    }

                    // Save the new document
                    try
                    {
                        var saveOpts = new SaveAsOptions { OverwriteExistingFile = overwrite };
                        newDoc.SaveAs(exportFilePath, saveOpts);
                        LogMessage($"Exported: {exportFileName} (workset: {worksetName})");
                        exportedCount++;
                    }
                    catch (Exception saveEx)
                    {
                        LogMessage($"ERROR saving workset file '{exportFileName}': {saveEx.Message}");
                    }
                }
                catch (Exception ex)
                {
                    LogMessage($"ERROR preparing export for workset '{worksetName}': {ex.Message}");
                }
                finally
                {
                    try
                    {
                        if (newDoc != null)
                        {
                            bool closeResult = newDoc.Close(false);
                            LogMessage($"Closed temporary document for workset '{worksetName}' (success: {closeResult})");
                        }
                    }
                    catch (Exception closeEx)
                    {
                        LogMessage($"Warning: Could not close temporary document for '{worksetName}': {closeEx.Message}");
                    }
                }
            }

            if (exportedCount == 0)
            {
                LogMessage("WARNING: No workset files were exported - no relevant elements found or all exports skipped.");
            }
            else
            {
                LogMessage($"Workset extraction completed: {exportedCount} files created.");
            }
        }

        /// <summary>
        /// Clean a string to be safe for file system use
        /// </summary>
        private string CleanFileSystemName(string input)
        {
            if (string.IsNullOrEmpty(input))
                return "Unknown";

            // Replace invalid file system characters
            char[] invalidChars = Path.GetInvalidFileNameChars();
            string cleaned = input;

            foreach (char c in invalidChars)
            {
                cleaned = cleaned.Replace(c, '_');
            }

            // Replace spaces and other problematic characters
            cleaned = cleaned.Replace(' ', '_')
                           .Replace('(', '_')
                           .Replace(')', '_')
                           .Replace('[', '_')
                           .Replace(']', '_')
                           .Replace('{', '_')
                           .Replace('}', '_');

            // Trim and ensure it's not empty
            cleaned = cleaned.Trim('_');
            if (string.IsNullOrEmpty(cleaned))
                cleaned = "Unknown";

            // Limit length
            if (cleaned.Length > 50)
                cleaned = cleaned.Substring(0, 50);

            return cleaned;
        }

        /// <summary>
        /// Template integration that preserves worksets in the exported files.
        /// NOTE: preserving worksets requires the saved file to be a workshared (central) file.
        /// </summary>
        public bool IntegrateIntoTemplate(List<string> extractedFiles, string templateFilePath, string destinationPath)
        {
            try
            {
                LogMessage($"Starting template integration with {extractedFiles?.Count ?? 0} files...");
                LogMessage($"Template file: {System.IO.Path.GetFileName(templateFilePath)}");

                if (!File.Exists(templateFilePath))
                {
                    LogMessage($"ERROR: Template file not found: {templateFilePath}");
                    return false;
                }

                // Create "In Template" subfolder
                string templateOutputPath = System.IO.Path.Combine(destinationPath, "In Template");
                if (!Directory.Exists(templateOutputPath))
                {
                    Directory.CreateDirectory(templateOutputPath);
                    LogMessage($"Created 'In Template' directory: {templateOutputPath}");
                }

                var app = _uiDoc.Application.Application;
                int processedCount = 0;

                foreach (var extractedFilePath in extractedFiles)
                {
                    LogMessage($"=== Processing file {processedCount + 1}/{extractedFiles.Count}: {System.IO.Path.GetFileName(extractedFilePath)} ===");

                    if (!File.Exists(extractedFilePath))
                    {
                        LogMessage($"ERROR: Extracted file not found: {extractedFilePath}. Skipping.");
                        continue;
                    }

                    string tempTemplatePath = null;
                    Document templateDoc = null;
                    Document extractedDoc = null;

                    try
                    {
                        // --- Create a fresh temporary copy of the template (per extracted file) ---
                        tempTemplatePath = System.IO.Path.Combine(destinationPath, $"temp_template_{processedCount:000}_{Guid.NewGuid():N}.rvt");
                        LogMessage($"Creating temporary template copy: {tempTemplatePath}");
                        File.Copy(templateFilePath, tempTemplatePath, true);
                        File.SetAttributes(tempTemplatePath, FileAttributes.Normal);

                        // --- Open the temporary template. Try to DETACH AND PRESERVE worksets so template keeps workset structure ---
                        LogMessage("Opening temporary template copy (attempting DetachAndPreserveWorksets)...");
                        try
                        {
                            OpenOptions openOpts = new OpenOptions();

                            try
                            {
                                // Prefer DetachAndPreserveWorksets so the opened doc keeps worksets
                                openOpts.DetachFromCentralOption = DetachFromCentralOption.DetachAndPreserveWorksets;
                            }
                            catch
                            {
                                // Some older API versions might not expose DetachFromCentralOption - ignore and fallback
                            }

                            ModelPath mp = null;
                            try
                            {
                                mp = ModelPathUtils.ConvertUserVisiblePathToModelPath(tempTemplatePath);
                            }
                            catch (Exception convEx)
                            {
                                LogMessage($"Warning: ModelPath conversion failed for '{tempTemplatePath}': {convEx.Message}");
                                mp = null;
                            }

                            if (mp != null)
                            {
                                templateDoc = app.OpenDocumentFile(mp, openOpts);
                            }
                            else
                            {
                                // fallback to simple open if overload not available
                                templateDoc = app.OpenDocumentFile(tempTemplatePath);
                            }

                            if (templateDoc == null)
                            {
                                LogMessage("Open returned null - falling back to simple OpenDocumentFile.");
                                templateDoc = app.OpenDocumentFile(tempTemplatePath);
                            }
                        }
                        catch (Exception openTempEx)
                        {
                            LogMessage($"WARNING: Opening temp template failed ({openTempEx.Message}). Trying regular open...");
                            try { templateDoc = app.OpenDocumentFile(tempTemplatePath); }
                            catch (Exception e2) { LogMessage($"ERROR opening temp template: {e2.Message}"); templateDoc = null; }
                        }

                        if (templateDoc == null)
                        {
                            LogMessage($"ERROR: Could not open temporary template copy: {tempTemplatePath}. Skipping file.");
                            try { if (File.Exists(tempTemplatePath)) File.Delete(tempTemplatePath); } catch { }
                            continue;
                        }

                        LogMessage($"Opened template copy: {templateDoc.Title} (IsWorkshared: {templateDoc.IsWorkshared})");

                        // --- Open the extracted file (read) ---
                        LogMessage("Opening extracted file for reading...");
                        try
                        {
                            extractedDoc = app.OpenDocumentFile(extractedFilePath);
                        }
                        catch (Exception openExtractEx)
                        {
                            LogMessage($"ERROR opening extracted file: {openExtractEx.Message}");
                            extractedDoc = null;
                        }

                        if (extractedDoc == null)
                        {
                            LogMessage($"ERROR: Could not open extracted file: {extractedFilePath}. Closing template and skipping.");
                            try { templateDoc.Close(false); } catch { }
                            try { if (File.Exists(tempTemplatePath)) File.Delete(tempTemplatePath); } catch { }
                            continue;
                        }

                        LogMessage($"Opened extracted file: {extractedDoc.Title}");

                        // --- Collect MEP elements from extracted file ---
                        var mepElements = GetAllMepElementsFromDocument(extractedDoc);
                        LogMessage($"Found {mepElements.Count} MEP elements in extracted file.");

                        if (mepElements.Count == 0)
                        {
                            LogMessage("No MEP elements found - closing docs and moving to next file.");
                            try { extractedDoc.Close(false); } catch { }
                            try { templateDoc.Close(false); } catch { }
                            try { if (File.Exists(tempTemplatePath)) File.Delete(tempTemplatePath); } catch { }
                            processedCount++;
                            continue;
                        }

                        // --- Copy MEP elements into the fresh templateDoc ---
                        var elementIds = mepElements.Select(e => e.Id).ToList();

                        LogMessage("Starting copy transaction on template copy...");
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

                                LogMessage($"Copied {copied?.Count ?? 0} elements into template copy.");
                                t.Commit();
                            }
                            catch (Exception copyEx)
                            {
                                t.RollBack();
                                LogMessage($"ERROR during copy: {copyEx.Message}");
                                try { extractedDoc.Close(false); } catch { }
                                try { templateDoc.Close(false); } catch { }
                                try { if (File.Exists(tempTemplatePath)) File.Delete(tempTemplatePath); } catch { }
                                continue;
                            }
                        }

                        // --- If the templateDoc is workshared we may want to synchronize/relinquish (optional) ---
                        if (templateDoc.IsWorkshared)
                        {
                            LogMessage("Template copy is workshared - attempting SyncWithCentral + relinquish (safe-wrapped)...");
                            try
                            {
                                var transOpts = new TransactWithCentralOptions();
                                transOpts.SetLockCallback(new SynchLockCallback());

                                var syncOpts = new SynchronizeWithCentralOptions();
                                var relOpts = new RelinquishOptions(true);
                                syncOpts.SetRelinquishOptions(relOpts);
                                syncOpts.SaveLocalAfter = false;

                                templateDoc.SynchronizeWithCentral(transOpts, syncOpts);
                                LogMessage("Synchronized and relinquished template copy.");
                            }
                            catch (Exception syncEx)
                            {
                                LogMessage($"WARNING: SyncWithCentral failed on template copy: {syncEx.Message} — continuing to save.");
                            }
                        }

                        // --- Save the integrated file to the output location (In Template folder).
                        // If templateDoc.IsWorkshared, we must set WorksharingSaveAsOptions.SaveAsCentral = true. ---
                        string outputFileName = System.IO.Path.GetFileName(extractedFilePath);
                        string outputFilePath = System.IO.Path.Combine(templateOutputPath, outputFileName);

                        LogMessage($"Saving integrated file as: {outputFilePath}");
                        try
                        {
                            var saveOpts = new SaveAsOptions { OverwriteExistingFile = true };

                            if (templateDoc.IsWorkshared)
                            {
                                try
                                {
                                    var wsOpts = new WorksharingSaveAsOptions { SaveAsCentral = true };
                                    saveOpts.SetWorksharingOptions(wsOpts);
                                    LogMessage("WorksharingSaveAsOptions.SaveAsCentral = true set for SaveAs (preserving worksets).");
                                }
                                catch (Exception exWs)
                                {
                                    LogMessage($"Warning: Could not set WorksharingSaveAsOptions: {exWs.Message}");
                                }
                            }

                            templateDoc.SaveAs(outputFilePath, saveOpts);
                            LogMessage($"Successfully saved integrated file: {outputFilePath}");
                        }
                        catch (Exception saveEx)
                        {
                            LogMessage($"ERROR saving integrated file: {saveEx.Message}");
                        }

                        // --- Close both docs ---
                        try { LogMessage("Closing extracted file..."); extractedDoc.Close(false); } catch (Exception cex) { LogMessage($"Warning closing extracted file: {cex.Message}"); }
                        try { LogMessage("Closing template copy..."); templateDoc.Close(false); } catch (Exception cex) { LogMessage($"Warning closing template copy: {cex.Message}"); }

                        // --- Remove temporary template file ---
                        try
                        {
                            if (!string.IsNullOrEmpty(tempTemplatePath) && File.Exists(tempTemplatePath))
                            {
                                File.SetAttributes(tempTemplatePath, FileAttributes.Normal);
                                File.Delete(tempTemplatePath);
                                LogMessage($"Deleted temporary template copy: {tempTemplatePath}");
                            }
                        }
                        catch (Exception exDelete) { LogMessage($"Warning deleting temp template file: {exDelete.Message}"); }

                        processedCount++;
                        LogMessage($"Successfully processed {processedCount}/{extractedFiles.Count}");
                    }
                    catch (Exception fileEx)
                    {
                        LogMessage($"ERROR processing file '{System.IO.Path.GetFileName(extractedFilePath)}': {fileEx.Message}");
                        LogMessage($"Stack trace: {fileEx.StackTrace}");
                        try { if (extractedDoc != null) extractedDoc.Close(false); } catch { }
                        try { if (templateDoc != null) templateDoc.Close(false); } catch { }
                        try { if (!string.IsNullOrEmpty(tempTemplatePath) && File.Exists(tempTemplatePath)) File.Delete(tempTemplatePath); } catch { }
                    }
                } // foreach

                LogMessage($"=== TEMPLATE INTEGRATION COMPLETED ===");
                LogMessage($"Successfully processed {processedCount} files.");
                LogMessage($"Output location: {templateOutputPath}");

                return true;
            }
            catch (Exception ex)
            {
                LogMessage($"ERROR in template integration: {ex.Message}");
                LogMessage($"Stack Trace: {ex.StackTrace}");
                return false;
            }
        }



        /// <summary>
        /// Public helper to safely open a document (reuses already-open docs and reports whether this method opened it)
        /// </summary>
        public Document SafeOpenDocument(Application app, string filePath, out bool openedByUs)
        {
            openedByUs = false;

            try
            {
                // Reuse document if already open in this Revit session
                var alreadyOpen = app.Documents.Cast<Document>()
                    .FirstOrDefault(d => string.Equals(d.PathName, filePath, StringComparison.OrdinalIgnoreCase));

                if (alreadyOpen != null)
                {
                    LogMessage($"File already open in session - reusing Document: {Path.GetFileName(filePath)}");
                    return alreadyOpen;
                }

                LogMessage($"Opening document via API: {filePath}");
                var doc = app.OpenDocumentFile(filePath);

                if (doc != null)
                {
                    openedByUs = true;
                    LogMessage($"Opened document: {doc.Title}");
                }
                else
                {
                    LogMessage($"OpenDocumentFile returned null for {filePath}");
                }

                return doc;
            }
            catch (Exception ex)
            {
                LogMessage($"ERROR opening document '{filePath}': {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// Gets the default 3D view (the house icon view) from a document
        /// </summary>
        private View3D GetDefault3DView(Document doc)
        {
            try
            {
                // Find the default {3D} view - this is typically the view that opens when you click the house icon
                var view3D = new FilteredElementCollector(doc)
                    .OfClass(typeof(View3D))
                    .Cast<View3D>()
                    .FirstOrDefault(v => v.Name == "{3D}" ||
                                        v.Name.Contains("3D") && v.IsTemplate == false);

                // If no {3D} view found, get the first non-template 3D view
                if (view3D == null)
                {
                    view3D = new FilteredElementCollector(doc)
                        .OfClass(typeof(View3D))
                        .Cast<View3D>()
                        .FirstOrDefault(v => !v.IsTemplate);
                }

                return view3D;
            }
            catch (Exception ex)
            {
                LogMessage($"Warning: Error finding 3D view: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// Gets all MEP elements from a specific document (not the main document)
        /// </summary>
        private List<Element> GetAllMepElementsFromDocument(Document doc)
        {
            var mepElements = new List<Element>();

            // Get all MEP-related elements (same categories as main workflow)
            var categories = new List<BuiltInCategory>
            {
                BuiltInCategory.OST_PipeFitting,
                BuiltInCategory.OST_PipeAccessory,
                BuiltInCategory.OST_PipeCurves,
                BuiltInCategory.OST_DuctFitting,
                BuiltInCategory.OST_DuctAccessory,
                BuiltInCategory.OST_DuctCurves,
                BuiltInCategory.OST_DuctTerminal,
                BuiltInCategory.OST_MechanicalEquipment,
                BuiltInCategory.OST_PlumbingFixtures,
                BuiltInCategory.OST_Sprinklers,
                BuiltInCategory.OST_ElectricalEquipment,
                BuiltInCategory.OST_ElectricalFixtures,
                BuiltInCategory.OST_LightingFixtures,
                BuiltInCategory.OST_CableTray,
                BuiltInCategory.OST_CableTrayFitting,
                BuiltInCategory.OST_Conduit,
                BuiltInCategory.OST_ConduitFitting
            };

            foreach (var category in categories)
            {
                try
                {
                    var collector = new FilteredElementCollector(doc)
                        .OfCategory(category)
                        .WhereElementIsNotElementType();

                    mepElements.AddRange(collector);
                }
                catch (Exception ex)
                {
                    LogMessage($"Warning: Could not collect category {category} from document: {ex.Message}");
                }
            }

            LogMessage($"Collected {mepElements.Count} total MEP elements from document '{doc.Title}'.");
            return mepElements;
        }

        private List<Element> GetAllMepElements()
        {
            var mepElements = new List<Element>();

            // Get all MEP-related elements
            var categories = new List<BuiltInCategory>
            {
                BuiltInCategory.OST_PipeFitting,
                BuiltInCategory.OST_PipeAccessory,
                BuiltInCategory.OST_PipeCurves,
                BuiltInCategory.OST_DuctFitting,
                BuiltInCategory.OST_DuctAccessory,
                BuiltInCategory.OST_DuctCurves,
                BuiltInCategory.OST_DuctTerminal,
                BuiltInCategory.OST_MechanicalEquipment,
                BuiltInCategory.OST_PlumbingFixtures,
                BuiltInCategory.OST_Sprinklers,
                BuiltInCategory.OST_ElectricalEquipment,
                BuiltInCategory.OST_ElectricalFixtures,
                BuiltInCategory.OST_LightingFixtures,
                BuiltInCategory.OST_CableTray,
                BuiltInCategory.OST_CableTrayFitting,
                BuiltInCategory.OST_Conduit,
                BuiltInCategory.OST_ConduitFitting
            };

            foreach (var category in categories)
            {
                try
                {
                    var collector = new FilteredElementCollector(_doc)
                        .OfCategory(category)
                        .WhereElementIsNotElementType();

                    mepElements.AddRange(collector);
                }
                catch (Exception ex)
                {
                    LogMessage($"Warning: Could not collect category {category}: {ex.Message}");
                }
            }

            LogMessage($"Collected {mepElements.Count} total MEP elements from {categories.Count} categories.");
            return mepElements;
        }

        private List<Element> FindMatchingElements(List<Element> elements, MappingRecord record)
        {
            var matchingElements = new List<Element>();
            string systemPattern = record.SystemNameInModel?.Trim();
            string systemDescription = record.SystemDescription?.Trim();

            LogMessage($"Searching for elements matching workset '{record.WorksetName}' with pattern '{systemPattern}' in {elements.Count} elements...");

            // Special handling for DX_ELT workset - check electrical elements first
            if (record.WorksetName?.ToUpper() == "DX_ELT")
            {
                LogMessage($"Processing DX_ELT workset - checking for electrical elements");

                foreach (var element in elements)
                {
                    if (IsElementElectricalType(element))
                    {
                        matchingElements.Add(element);
                    }
                }

                LogMessage($"Found {matchingElements.Count} electrical elements for DX_ELT");
                return matchingElements;
            }

            // Special handling for records with empty/null system patterns
            if (string.IsNullOrEmpty(systemPattern) || systemPattern == "-")
            {
                LogMessage($"No system pattern provided for '{record.WorksetName}' - using category and parameter-based matching");

                foreach (var element in elements)
                {
                    if (IsElementMatchingByCategory(element, record))
                    {
                        matchingElements.Add(element);
                    }
                }

                return matchingElements;
            }

            // Standard pattern-based matching for elements with system names
            foreach (var element in elements)
            {
                if (IsElementMatchingSystem(element, systemPattern, systemDescription))
                {
                    matchingElements.Add(element);
                }
            }

            return matchingElements;
        }

        private bool IsElementMatchingByCategory(Element element, MappingRecord record)
        {
            try
            {
                string worksetName = record.WorksetName?.ToUpper();
                string description = record.SystemDescription?.ToUpper();

                // Handle DX_ELT - Electrical elements (cable trays, busbars)
                if (worksetName == "DX_ELT")
                {
                    if (IsElementElectricalType(element))
                    {
                        LogMessage($"  MATCH (Category-ELT): Element {element.Id} identified as electrical");
                        return true;
                    }
                }

                // Handle DX_STB - Structural elements
                else if (worksetName == "DX_STB")
                {
                    if (IsElementStructuralType(element))
                    {
                        LogMessage($"  MATCH (Category-STB): Element {element.Id} identified as structural");
                        return true;
                    }
                }

                // Handle DX_RR - Cleanroom Partitions
                else if (worksetName == "DX_RR")
                {
                    if (IsElementCleanroomPartition(element))
                    {
                        LogMessage($"  MATCH (Category-RR): Element {element.Id} identified as cleanroom partition");
                        return true;
                    }
                }

                // Handle DX_FND - Process Tool Pedestal
                else if (worksetName == "DX_FND")
                {
                    if (IsElementFoundation(element))
                    {
                        LogMessage($"  MATCH (Category-FND): Element {element.Id} identified as foundation/pedestal");
                        return true;
                    }
                }

                // Generic parameter-based matching for other categories
                else
                {
                    if (HasMatchingWorksetParameter(element, worksetName, description))
                    {
                        LogMessage($"  MATCH (Parameter): Element {element.Id} matched by workset parameter");
                        return true;
                    }
                }

                return false;
            }
            catch (Exception ex)
            {
                LogMessage($"Warning: Error in category matching for element {element.Id}: {ex.Message}");
                return false;
            }
        }



        private bool IsElementStructuralType(Element element)
        {
            try
            {
                // Check structural categories
                var structuralCategories = new[]
                {
            BuiltInCategory.OST_StructuralFraming,
            BuiltInCategory.OST_StructuralColumns,
            BuiltInCategory.OST_StructuralFoundation,
            BuiltInCategory.OST_StructuralFramingSystem,
            BuiltInCategory.OST_StructuralStiffener,
            BuiltInCategory.OST_StructuralTruss
        };

                var elementCategory = element.Category;
                if (elementCategory != null)
                {
                    var categoryId = (BuiltInCategory)elementCategory.Id.IntegerValue;
                    if (structuralCategories.Contains(categoryId))
                    {
                        LogMessage($"  Structural element {element.Id} - category match: {categoryId}");
                        return true;
                    }
                }

                return false;
            }
            catch (Exception ex)
            {
                LogMessage($"Warning: Error checking structural type for element {element.Id}: {ex.Message}");
                return false;
            }
        }

        private bool IsElementCleanroomPartition(Element element)
        {
            try
            {
                // Check for walls, generic models, or family instances that could be partitions
                var partitionCategories = new[]
                {
            BuiltInCategory.OST_Walls,
            BuiltInCategory.OST_GenericModel,
            BuiltInCategory.OST_Doors,
            BuiltInCategory.OST_Windows
        };

                var elementCategory = element.Category;
                if (elementCategory != null)
                {
                    var categoryId = (BuiltInCategory)elementCategory.Id.IntegerValue;
                    if (partitionCategories.Contains(categoryId))
                    {
                        return CheckCleanroomWorksetParameter(element);
                    }
                }

                return false;
            }
            catch (Exception ex)
            {
                LogMessage($"Warning: Error checking cleanroom partition type for element {element.Id}: {ex.Message}");
                return false;
            }
        }

        private bool IsElementFoundation(Element element)
        {
            try
            {
                // Check for foundation-related categories
                var foundationCategories = new[]
                {
            BuiltInCategory.OST_StructuralFoundation,
            BuiltInCategory.OST_GenericModel,
            BuiltInCategory.OST_MechanicalEquipment // Tool pedestals might be MEP equipment
        };

                var elementCategory = element.Category;
                if (elementCategory != null)
                {
                    var categoryId = (BuiltInCategory)elementCategory.Id.IntegerValue;
                    if (foundationCategories.Contains(categoryId))
                    {
                        return CheckFoundationWorksetParameter(element);
                    }
                }

                return false;
            }
            catch (Exception ex)
            {
                LogMessage($"Warning: Error checking foundation type for element {element.Id}: {ex.Message}");
                return false;
            }
        }

        private bool CheckStructuralWorksetParameter(Element element)
        {
            try
            {
                // Get the element's type
                ElementType elementType = element.Document.GetElement(element.GetTypeId()) as ElementType;

                // Check type parameters first
                if (elementType != null)
                {
                    var worksetParam = elementType.LookupParameter("Workset");
                    if (worksetParam != null && !string.IsNullOrEmpty(worksetParam.AsString()))
                    {
                        string worksetValue = worksetParam.AsString();
                        LogMessage($"  Structural element {element.Id} has Workset parameter: '{worksetValue}'");

                        var structuralKeywords = new[]
                        {
                    "Steel", "Structure", "Structural", "STB", "Frame", "Column", "Beam"
                };

                        bool isStructural = structuralKeywords.Any(keyword =>
                            worksetValue.IndexOf(keyword, StringComparison.OrdinalIgnoreCase) >= 0);

                        return isStructural;
                    }
                }

                // Check instance parameters
                var instanceWorksetParam = element.LookupParameter("Workset");
                if (instanceWorksetParam != null && !string.IsNullOrEmpty(instanceWorksetParam.AsString()))
                {
                    string worksetValue = instanceWorksetParam.AsString();
                    var structuralKeywords = new[] { "Steel", "Structure", "Structural", "STB" };

                    return structuralKeywords.Any(keyword =>
                        worksetValue.IndexOf(keyword, StringComparison.OrdinalIgnoreCase) >= 0);
                }

                // Default to true for structural categories
                var elementCategory = element.Category;
                if (elementCategory != null)
                {
                    var categoryId = (BuiltInCategory)elementCategory.Id.IntegerValue;
                    var pureStructuralCategories = new[]
                    {
                BuiltInCategory.OST_StructuralFraming,
                BuiltInCategory.OST_StructuralColumns,
                BuiltInCategory.OST_StructuralFoundation
            };

                    if (pureStructuralCategories.Contains(categoryId))
                    {
                        LogMessage($"  Structural element {element.Id} - pure structural category, defaulting to true");
                        return true;
                    }
                }

                return false;
            }
            catch (Exception ex)
            {
                LogMessage($"Warning: Error checking structural workset parameter for element {element.Id}: {ex.Message}");
                return false;
            }
        }

        private bool CheckCleanroomWorksetParameter(Element element)
        {
            try
            {
                // Check both type and instance parameters
                var parameterSources = new List<Element> { element };

                ElementType elementType = element.Document.GetElement(element.GetTypeId()) as ElementType;
                if (elementType != null)
                    parameterSources.Add(elementType);

                foreach (var paramSource in parameterSources)
                {
                    var worksetParam = paramSource.LookupParameter("Workset");
                    if (worksetParam != null && !string.IsNullOrEmpty(worksetParam.AsString()))
                    {
                        string worksetValue = worksetParam.AsString();
                        var cleanroomKeywords = new[]
                        {
                    "Cleanroom", "Partition", "RR", "Clean Room", "Wall"
                };

                        bool isCleanroom = cleanroomKeywords.Any(keyword =>
                            worksetValue.IndexOf(keyword, StringComparison.OrdinalIgnoreCase) >= 0);

                        if (isCleanroom)
                        {
                            LogMessage($"  Cleanroom element {element.Id} matched workset parameter: '{worksetValue}'");
                            return true;
                        }
                    }
                }

                return false;
            }
            catch (Exception ex)
            {
                LogMessage($"Warning: Error checking cleanroom workset parameter for element {element.Id}: {ex.Message}");
                return false;
            }
        }

        private bool CheckFoundationWorksetParameter(Element element)
        {
            try
            {
                // Check both type and instance parameters
                var parameterSources = new List<Element> { element };

                ElementType elementType = element.Document.GetElement(element.GetTypeId()) as ElementType;
                if (elementType != null)
                    parameterSources.Add(elementType);

                foreach (var paramSource in parameterSources)
                {
                    var worksetParam = paramSource.LookupParameter("Workset");
                    if (worksetParam != null && !string.IsNullOrEmpty(worksetParam.AsString()))
                    {
                        string worksetValue = worksetParam.AsString();
                        var foundationKeywords = new[]
                        {
                    "Foundation", "Pedestal", "FND", "Tool", "Base"
                };

                        bool isFoundation = foundationKeywords.Any(keyword =>
                            worksetValue.IndexOf(keyword, StringComparison.OrdinalIgnoreCase) >= 0);

                        if (isFoundation)
                        {
                            LogMessage($"  Foundation element {element.Id} matched workset parameter: '{worksetValue}'");
                            return true;
                        }
                    }
                }

                return false;
            }
            catch (Exception ex)
            {
                LogMessage($"Warning: Error checking foundation workset parameter for element {element.Id}: {ex.Message}");
                return false;
            }
        }

        private bool HasMatchingWorksetParameter(Element element, string targetWorksetName, string description)
        {
            try
            {
                // Check both type and instance parameters
                var parameterSources = new List<Element> { element };

                ElementType elementType = element.Document.GetElement(element.GetTypeId()) as ElementType;
                if (elementType != null)
                    parameterSources.Add(elementType);

                foreach (var paramSource in parameterSources)
                {
                    var worksetParam = paramSource.LookupParameter("Workset");
                    if (worksetParam != null && !string.IsNullOrEmpty(worksetParam.AsString()))
                    {
                        string worksetValue = worksetParam.AsString();

                        // Check if workset parameter contains target workset name (without DX_ prefix)
                        string cleanTargetName = targetWorksetName.Replace("DX_", "");
                        if (worksetValue.IndexOf(cleanTargetName, StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            LogMessage($"  Element {element.Id} matched by workset parameter: '{worksetValue}' contains '{cleanTargetName}'");
                            return true;
                        }

                        // Check against description keywords if provided
                        if (!string.IsNullOrEmpty(description))
                        {
                            var descriptionWords = description.Split(new char[] { ' ', ',', ';' }, StringSplitOptions.RemoveEmptyEntries);
                            foreach (var word in descriptionWords)
                            {
                                if (word.Length > 3 && worksetValue.IndexOf(word, StringComparison.OrdinalIgnoreCase) >= 0)
                                {
                                    LogMessage($"  Element {element.Id} matched by description word: '{word}' in workset parameter: '{worksetValue}'");
                                    return true;
                                }
                            }
                        }
                    }
                }

                return false;
            }
            catch (Exception ex)
            {
                LogMessage($"Warning: Error checking workset parameter for element {element.Id}: {ex.Message}");
                return false;
            }
        }

        private bool IsElementMatchingSystem(Element element, string systemPattern, string systemDescription)
        {
            try
            {
                // Special handling for ELT (Electrical) elements that don't have system names
                if (IsElectricalPattern(systemPattern))
                {
                    if (IsElementElectricalType(element))
                    {
                        LogMessage($"  MATCH (Electrical): Element {element.Id} - Cable Tray/Busbar detected for ELT");
                        return true;
                    }
                }

                // Get all possible system-related parameters (existing logic)
                var systemValues = new List<string>();

                // Try different parameter approaches for different element types
                var parameterIds = new List<BuiltInParameter>
        {
            BuiltInParameter.RBS_SYSTEM_NAME_PARAM,           // System Name
            BuiltInParameter.RBS_SYSTEM_CLASSIFICATION_PARAM, // System Classification  
            BuiltInParameter.RBS_SYSTEM_ABBREVIATION_PARAM,   // System Abbreviation
            BuiltInParameter.ELEM_TYPE_PARAM,                 // Element Type
        };

                foreach (var paramId in parameterIds)
                {
                    var param = element.get_Parameter(paramId);
                    if (param != null && !string.IsNullOrEmpty(param.AsString()))
                    {
                        systemValues.Add(param.AsString());
                    }
                }

                // For MEP elements, also try to get system information from ConnectorManager
                if (element is MEPCurve mepCurve)
                {
                    try
                    {
                        var systemParam = mepCurve.get_Parameter(BuiltInParameter.RBS_SYSTEM_NAME_PARAM);
                        if (systemParam != null && !string.IsNullOrEmpty(systemParam.AsString()))
                        {
                            systemValues.Add(systemParam.AsString());
                        }
                    }
                    catch { }
                }

                // For Family Instances, check additional parameters
                if (element is FamilyInstance familyInstance)
                {
                    try
                    {
                        // Try to get system from MEP model
                        if (familyInstance.MEPModel != null)
                        {
                            var sysNameParam = familyInstance.get_Parameter(BuiltInParameter.RBS_SYSTEM_NAME_PARAM);
                            if (sysNameParam != null && !string.IsNullOrEmpty(sysNameParam.AsString()))
                            {
                                systemValues.Add(sysNameParam.AsString());
                            }
                        }
                    }
                    catch { }
                }

                // Remove duplicates and empty values
                systemValues = systemValues.Where(v => !string.IsNullOrEmpty(v)).Distinct().ToList();

                if (systemValues.Count == 0)
                    return false;

                // Check if any system value matches our pattern
                foreach (var systemValue in systemValues)
                {
                    if (DoesSystemValueMatchPattern(systemValue, systemPattern))
                    {
                        LogMessage($"  MATCH: Element {element.Id} matched pattern.");
                        return true;
                    }
                }

                // Fallback: check system description if provided
                if (!string.IsNullOrEmpty(systemDescription))
                {
                    foreach (var systemValue in systemValues)
                    {
                        if (systemValue.IndexOf(systemDescription, StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            LogMessage($"  MATCH (Description): Element {element.Id} - contains description.");
                            return true;
                        }
                    }
                }

                return false;
            }
            catch (Exception ex)
            {
                LogMessage($"Warning: Error checking element {element.Id}: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// Checks if the pattern indicates electrical elements (ELT, Electrical, etc.)
        /// </summary>
        private bool IsElectricalPattern(string pattern)
        {
            if (string.IsNullOrEmpty(pattern))
                return false;

            var electricalKeywords = new[] { "ELT", "ELECTRICAL", "CABLE", "BUSBAR", "POWER", "LIGHTING" };

            return electricalKeywords.Any(keyword =>
                pattern.IndexOf(keyword, StringComparison.OrdinalIgnoreCase) >= 0);
        }

        /// <summary>
        /// Determines if an element is an electrical type (Cable Tray, Busbar, etc.)
        /// that should go to ELT workset
        /// </summary>
        private bool IsElementElectricalType(Element element)
        {
            try
            {
                // Check if element is in electrical categories
                var electricalCategories = new[]
                {
            BuiltInCategory.OST_CableTray,
            BuiltInCategory.OST_CableTrayFitting,
            BuiltInCategory.OST_Conduit,
            BuiltInCategory.OST_ConduitFitting,
            BuiltInCategory.OST_ElectricalEquipment,
            BuiltInCategory.OST_ElectricalFixtures,
            BuiltInCategory.OST_LightingFixtures
        };

                var elementCategory = element.Category;
                if (elementCategory != null)
                {
                    var categoryId = (BuiltInCategory)elementCategory.Id.IntegerValue;
                    if (electricalCategories.Contains(categoryId))
                    {
                        // Additional check for cable trays using the Workset type parameter
                        if (categoryId == BuiltInCategory.OST_CableTray ||
                            categoryId == BuiltInCategory.OST_CableTrayFitting)
                        {
                            return CheckCableTrayWorksetParameter(element);
                        }

                        // For other electrical elements, include them by default
                        return true;
                    }
                }

                return false;
            }
            catch (Exception ex)
            {
                LogMessage($"Warning: Error checking electrical type for element {element.Id}: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// Checks if a Cable Tray element has the correct Workset type parameter
        /// </summary>
        private bool CheckCableTrayWorksetParameter(Element element)
        {
            try
            {
                // Get the element's type
                ElementType elementType = element.Document.GetElement(element.GetTypeId()) as ElementType;
                if (elementType == null)
                    return true; // Default to electrical if we can't get the type

                // Look for "Workset" parameter in type parameters
                var worksetParam = elementType.LookupParameter("Workset");
                if (worksetParam != null && !string.IsNullOrEmpty(worksetParam.AsString()))
                {
                    string worksetValue = worksetParam.AsString();
                    LogMessage($"  Cable Tray/Fitting {element.Id} has Workset parameter: '{worksetValue}'");

                    // Check if it explicitly excludes electrical (rare case)
                    var nonElectricalKeywords = new[] { "Fire", "HVAC", "Plumbing", "NOT ELECTRICAL" };

                    bool isNonElectrical = nonElectricalKeywords.Any(keyword =>
                        worksetValue.IndexOf(keyword, StringComparison.OrdinalIgnoreCase) >= 0);

                    // If explicitly non-electrical, return false; otherwise assume electrical
                    return !isNonElectrical;
                }

                // If no Workset parameter found, default to true for cable trays
                LogMessage($"  Cable Tray/Fitting {element.Id} - no Workset parameter found, defaulting to electrical");
                return true;
            }
            catch (Exception ex)
            {
                LogMessage($"Warning: Error checking cable tray workset parameter for element {element.Id}: {ex.Message}");
                return true; // Default to electrical if we can't determine
            }
        }

        /// <summary>
        /// Checks family instance for electrical workset parameters
        /// </summary>
        private bool CheckFamilyInstanceWorksetParameter(FamilyInstance familyInstance)
        {
            try
            {
                // Check instance parameters first
                var worksetParam = familyInstance.LookupParameter("Workset");
                if (worksetParam != null && !string.IsNullOrEmpty(worksetParam.AsString()))
                {
                    string worksetValue = worksetParam.AsString();
                    var electricalKeywords = new[] { "Cable", "Electrical", "Power", "ELT", "Busbar" };

                    return electricalKeywords.Any(keyword =>
                        worksetValue.IndexOf(keyword, StringComparison.OrdinalIgnoreCase) >= 0);
                }

                // Check type parameters
                ElementType elementType = familyInstance.Document.GetElement(familyInstance.GetTypeId()) as ElementType;
                if (elementType != null)
                {
                    var typeWorksetParam = elementType.LookupParameter("Workset");
                    if (typeWorksetParam != null && !string.IsNullOrEmpty(typeWorksetParam.AsString()))
                    {
                        string worksetValue = typeWorksetParam.AsString();
                        var electricalKeywords = new[] { "Cable", "Electrical", "Power", "ELT", "Busbar" };

                        return electricalKeywords.Any(keyword =>
                            worksetValue.IndexOf(keyword, StringComparison.OrdinalIgnoreCase) >= 0);
                    }
                }

                return false;
            }
            catch (Exception ex)
            {
                LogMessage($"Warning: Error checking family instance workset parameter for element {familyInstance.Id}: {ex.Message}");
                return false;
            }
        }

        private bool DoesSystemValueMatchPattern(string systemValue, string pattern)
        {
            if (string.IsNullOrEmpty(systemValue) || string.IsNullOrEmpty(pattern))
                return false;

            // Handle exact match first
            if (systemValue.Equals(pattern, StringComparison.OrdinalIgnoreCase))
                return true;

            try
            {
                // Convert pattern to regex (replace 'x' with digits, handle other wildcards)
                string regexPattern = Regex.Escape(pattern)
                    .Replace("xxx", @"\d{2,3}")  // xxx becomes 2-3 digits
                    .Replace("xx", @"\d{1,3}")   // xx becomes 1-3 digits  
                    .Replace("x", @"\d")         // x becomes single digit
                    .Replace(@"\*", ".*");       // escaped * becomes wildcard

                var regex = new Regex(regexPattern, RegexOptions.IgnoreCase);

                if (regex.IsMatch(systemValue))
                    return true;

                // Additional check: see if the pattern (without x's) is contained in the system value
                string simplifiedPattern = pattern.Replace("xxx", "").Replace("xx", "").Replace("x", "").Trim();
                if (!string.IsNullOrEmpty(simplifiedPattern) &&
                    systemValue.IndexOf(simplifiedPattern, StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    return true;
                }
            }
            catch (Exception ex)
            {
                LogMessage($"Regex error for pattern '{pattern}': {ex.Message}");
            }

            // Final fallback: simple contains check
            string basePattern = pattern.Replace("xxx", "").Replace("xx", "").Replace("x", "").Trim();
            return !string.IsNullOrEmpty(basePattern) &&
                   systemValue.IndexOf(basePattern, StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private WorksetId GetOrCreateWorkset(string worksetName)
        {
            Workset workset = new FilteredWorksetCollector(_doc)
                .OfKind(WorksetKind.UserWorkset)
                .FirstOrDefault(w => w.Name.Equals(worksetName, StringComparison.CurrentCultureIgnoreCase));

            if (workset != null)
                return workset.Id;

            try
            {
                var newWorkset = Workset.Create(_doc, worksetName);
                LogMessage($"Created new workset: {worksetName}");
                return newWorkset.Id;
            }
            catch (Exception ex)
            {
                LogMessage($"ERROR creating workset '{worksetName}': {ex.Message}");
                return WorksetId.InvalidWorksetId;
            }
        }

        /// <summary>
        /// New export flow:
        ///  - Sync with central and relinquish ownership to avoid borrow issues
        ///  - For each package group, create a fresh new project document, copy only the element ids
        ///    (ElementTransformUtils.CopyElements) into the new document inside a transaction
        ///  - SaveAs the new document and close it
        /// </summary>
        private void ExportRVTs(Dictionary<string, List<ElementId>> packageGroups, string destinationPath, bool overwrite, bool exportQc)
        {
            string projectPrefix = System.IO.Path.GetFileNameWithoutExtension(_doc.PathName);
            if (projectPrefix.Contains('_'))
                projectPrefix = projectPrefix.Split('_')[0];

            // Add QC elements if requested and they exist
            if (exportQc)
            {
                var qcWorkset = new FilteredWorksetCollector(_doc)
                    .OfKind(WorksetKind.UserWorkset)
                    .FirstOrDefault(w => w.Name.Equals("DX_QC", StringComparison.CurrentCultureIgnoreCase));

                if (qcWorkset != null)
                {
                    var qcElements = new FilteredElementCollector(_doc)
                        .WhereElementIsNotElementType()
                        .Where(e => e.WorksetId == qcWorkset.Id)
                        .Select(e => e.Id)
                        .ToList();

                    if (qcElements.Any())
                    {
                        packageGroups["QC"] = qcElements;
                        LogMessage($"Added {qcElements.Count} elements to QC export group.");
                    }
                    else
                    {
                        LogMessage("No elements found in DX_QC workset - skipping QC export.");
                    }
                }
            }

            // Remove empty groups and log them
            var emptyGroups = packageGroups.Where(g => g.Value == null || !g.Value.Any()).Select(g => g.Key).ToList();
            foreach (var emptyKey in emptyGroups)
            {
                packageGroups.Remove(emptyKey);
                LogMessage($"Skipping export for '{emptyKey}' - no elements found.");
            }

            LogMessage($"Preparing to export {packageGroups.Count} model files...");

            // 1) Synchronize with central and relinquish ownership (if this is a workshared model)
            if (_doc.IsWorkshared)
            {
                try
                {
                    LogMessage("Attempting to synchronize with central and relinquish ownership...");

                    // Build TransactWithCentralOptions and SynchronizeWithCentralOptions
                    var transOpts = new TransactWithCentralOptions();
                    transOpts.SetLockCallback(new SynchLockCallback()); // if central is locked, we will give up quickly rather than waiting

                    var syncOpts = new SynchronizeWithCentralOptions();

                    // RelinquishOptions: true = relinquish owned elements/worksets after sync.
                    // If you want to relinquish selectively, construct RelinquishOptions appropriately.
                    var relinquishOpts = new RelinquishOptions(true);
                    syncOpts.SetRelinquishOptions(relinquishOpts);

                    // Do not save local after sync (optional; set to true if you want a local save)
                    syncOpts.SaveLocalAfter = false;

                    // Call synchronize (two-argument overload)
                    _doc.SynchronizeWithCentral(transOpts, syncOpts);

                    LogMessage("Synchronized with central and relinquished ownership (per options).");
                }
                catch (Exception ex)
                {
                    LogMessage($"WARNING: Synchronization with central failed: {ex.Message}");
                    // Proceeding anyway — copying only selected elements below may still succeed.
                }
            }
            else
            {
                LogMessage("Document is not workshared - skipping synchronize/relinquish step.");
            }

            // 2) For each package group, create a new project document and copy elements into it
            int exportedCount = 0;
            foreach (var group in packageGroups)
            {
                if (string.Equals(group.Key, "NO EXPORT", StringComparison.OrdinalIgnoreCase))
                    continue;

                string exportFileName = $"{projectPrefix}_{group.Key}_MO_Part_001_DX.rvt";
                string exportFilePath = System.IO.Path.Combine(destinationPath, exportFileName);

                // If file exists and we should not overwrite, skip.
                if (System.IO.File.Exists(exportFilePath) && !overwrite)
                {
                    LogMessage($"Skipped - File exists and overwrite is false: {exportFileName}");
                    continue;
                }

                // Create new project document using default template if available
                Document newDoc = null;
                try
                {
                    string templatePath = null;
                    try
                    {
                        // Prefer Revit's DefaultProjectTemplate if available
                        templatePath = _uiDoc.Application.Application.DefaultProjectTemplate;
                    }
                    catch
                    {
                        templatePath = null;
                    }

                    if (string.IsNullOrEmpty(templatePath) || !File.Exists(templatePath))
                    {
                        // Fallback: create new project using UnitSystem overload (create a new default doc)
                        LogMessage("Default project template not available or not found - creating new blank project document.");
                        newDoc = _uiDoc.Application.Application.NewProjectDocument(UnitSystem.Metric); // choose UnitSystem.Metric as default; change if needed
                    }
                    else
                    {
                        newDoc = _uiDoc.Application.Application.NewProjectDocument(templatePath);
                    }

                    if (newDoc == null)
                        throw new Exception("Failed to create new project document.");

                    LogMessage($"Created temporary new project document for export '{group.Key}'.");

                    // Prepare the list of element ids to copy: ensure distinct and still valid in source doc
                    var sourceElementIds = group.Value.Where(id => id != null).Distinct().ToList();

                    if (!sourceElementIds.Any())
                    {
                        LogMessage($"No valid elements to copy for group '{group.Key}' - skipping.");
                        // Close the temporary doc cleanly
                        try { newDoc.Close(false); } catch { }
                        continue;
                    }

                    // Copy elements from current document into the new document inside a transaction on the destination
                    try
                    {
                        using (Transaction tNew = new Transaction(newDoc, $"Copy elements for {group.Key}"))
                        {
                            tNew.Start();

                            // Use ElementTransformUtils.CopyElements to copy from source to destination
                            var copyOptions = new CopyPasteOptions();
                            // Optionally, configure copyOptions (duplicate type handling, etc.)

                            ICollection<ElementId> copiedIds = ElementTransformUtils.CopyElements(
                                _doc,
                                sourceElementIds,
                                newDoc,
                                Transform.Identity,
                                copyOptions
                            );

                            tNew.Commit();

                            LogMessage($"Copied {copiedIds?.Count ?? 0} elements into temporary document for group '{group.Key}'.");
                        }
                    }
                    catch (Exception copyEx)
                    {
                        LogMessage($"ERROR copying elements for group '{group.Key}': {copyEx.Message}");
                        // Try to close the temporary doc and continue
                        try { newDoc.Close(false); } catch { }
                        continue;
                    }

                    // Save the new document to disk
                    try
                    {
                        var saveOpts = new SaveAsOptions
                        {
                            OverwriteExistingFile = overwrite
                        };

                        // If you want to make the saved file workshared (central/local) change SaveAsOptions accordingly.
                        newDoc.SaveAs(exportFilePath, saveOpts);
                        LogMessage($"Exported: {exportFileName} with {group.Value.Count} elements");
                        exportedCount++;
                    }
                    catch (Exception saveEx)
                    {
                        LogMessage($"ERROR saving exported file '{exportFileName}': {saveEx.Message}");
                    }
                }
                catch (Exception ex)
                {
                    LogMessage($"ERROR preparing export for group '{group.Key}': {ex.Message}");
                }
                finally
                {
                    // Close the temporary document. Close(true) would attempt to save; but we already saved.
                    try
                    {
                        if (newDoc != null)
                        {
                            // Close without saving (we already saved with SaveAs). Use Close(false) to avoid prompts.
                            bool closeResult = newDoc.Close(false);
                            LogMessage($"Temporary document closed for group '{group.Key}' (success: {closeResult}).");
                        }
                    }
                    catch (Exception closeEx)
                    {
                        LogMessage($"Warning: could not close temporary document for '{group.Key}': {closeEx.Message}");
                    }
                }
            }

            if (exportedCount == 0)
            {
                LogMessage("WARNING: No model files were exported - no matching elements found or all exports skipped.");
            }
            else
            {
                LogMessage($"Completed exports: {exportedCount} files created.");
            }
        }

        // A small callback implementation for central lock handling. Returning false means "don't wait, give up".
        private class SynchLockCallback : ICentralLockedCallback
        {
            /// <summary>
            /// Return true to wait/retry for central lock, false to give up immediately.
            /// We return false here so the sync will fail fast if central is locked.
            /// </summary>
            /// <returns></returns>
            public bool ShouldWaitForLockAvailability()
            {
                return false;
            }
        }
    }
}
