using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

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
                        LogMessage($"  MATCH: Element {element.Id} - '{systemValue}' matches '{systemPattern}'");
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
                            LogMessage($"  MATCH (Description): Element {element.Id} - '{systemValue}' contains '{systemDescription}'");
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

                // Don't use word boundaries as they can be too restrictive for system names
                // Instead, just make it case insensitive
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
            var emptyGroups = packageGroups.Where(g => g.Value == null || !g.Value.Any()).ToList();
            foreach (var emptyGroup in emptyGroups)
            {
                packageGroups.Remove(emptyGroup.Key);
                LogMessage($"Skipping export for '{emptyGroup.Key}' - no elements found.");
            }

            LogMessage($"Exporting {packageGroups.Count} model files...");

            foreach (var group in packageGroups)
            {
                if (group.Key == "NO EXPORT") continue;

                string exportFileName = $"{projectPrefix}_{group.Key}_MO_Part_001_DX.rvt";
                string exportFilePath = System.IO.Path.Combine(destinationPath, exportFileName);

                if (System.IO.File.Exists(exportFilePath) && !overwrite)
                {
                    LogMessage($"Skipped - File exists and overwrite is false: {exportFileName}");
                    continue;
                }

                try
                {
                    SaveAsOptions options = new SaveAsOptions
                    {
                        OverwriteExistingFile = overwrite
                    };
                    _doc.SaveAs(exportFilePath, options);
                    LogMessage($"Exported: {exportFileName} with {group.Value.Count} elements");
                }
                catch (Exception ex)
                {
                    LogMessage($"ERROR exporting {exportFileName}: {ex.Message}");
                }
            }

            if (packageGroups.Count == 0)
            {
                LogMessage("WARNING: No model files were exported - no matching elements found.");
            }
        }
    }
}