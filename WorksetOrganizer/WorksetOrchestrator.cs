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

            if (string.IsNullOrEmpty(systemPattern) || systemPattern == "-")
                return matchingElements;

            LogMessage($"Searching for pattern '{systemPattern}' in {elements.Count} elements...");

            foreach (var element in elements)
            {
                if (IsElementMatchingSystem(element, systemPattern, systemDescription))
                {
                    matchingElements.Add(element);
                }
            }

            return matchingElements;
        }

        private bool IsElementMatchingSystem(Element element, string systemPattern, string systemDescription)
        {
            try
            {
                // Get all possible system-related parameters
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
                        // Note: user removed detailed MATCH lines from log; keep a short message
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
