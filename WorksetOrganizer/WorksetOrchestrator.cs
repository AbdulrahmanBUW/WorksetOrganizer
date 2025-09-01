using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Events;
using Autodesk.Revit.UI;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using RevitParameter = Autodesk.Revit.DB.Parameter;


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

        private void LogMessage(string message)
        {
            string formattedMessage = $"{DateTime.Now:HH:mm:ss} - {message}";
            _log.AppendLine(formattedMessage);
            LogUpdated?.Invoke(this, formattedMessage);
        }

        public bool Execute(List<MappingRecord> mapping, string destinationPath, bool overwriteFiles, bool exportQc)
        {
            try
            {
                // 1. VERIFY WORKSHARING
                if (!_doc.IsWorkshared)
                {
                    LogMessage("ERROR: Document is not workshared. Please enable worksharing.");
                    return false;
                }

                // 2. GET/CREATE THE DX_QC WORKSET
                WorksetId qcWorksetId = GetOrCreateWorkset("DX_QC");

                using (Transaction trans = new Transaction(_doc, "Process Workset Mapping"))
                {
                    trans.Start();

                    // 3. PROCESS EACH MAPPING RECORD
                    var packageGroupMapping = new Dictionary<string, List<ElementId>>();

                    foreach (var record in mapping)
                    {
                        LogMessage($"Processing record for system: {record.SystemNameInModel}");

                        // Skip processing for NO EXPORT
                        if (record.ModelPackageCode == "NO EXPORT")
                        {
                            LogMessage($"Skipped - Marked as 'NO EXPORT'.");
                            continue;
                        }

                        // Get or create the target workset for this record
                        WorksetId targetWorksetId = GetOrCreateWorkset(record.WorksetName);

                        // Find all elements with the matching System Name parameter
                        FilteredElementCollector collector = new FilteredElementCollector(_doc);
                        var elements = collector
                            .OfClass(typeof(FamilyInstance))
                            .WhereElementIsNotElementType()
                            .Where(e =>
                                e.get_Parameter(BuiltInParameter.RBS_SYSTEM_NAME_PARAM) != null &&
                                e.get_Parameter(BuiltInParameter.RBS_SYSTEM_NAME_PARAM).AsString() == record.SystemNameInModel
                            )
                            .ToList();

                        LogMessage($"Found {elements.Count} elements for system '{record.SystemNameInModel}'.");

                        // Move elements to the target workset
                        int movedCount = 0;
                        foreach (Element element in elements)
                        {
                            RevitParameter worksetParam = element.get_Parameter(BuiltInParameter.ELEM_PARTITION_PARAM);
                            if (worksetParam != null && !worksetParam.IsReadOnly && element.WorksetId != targetWorksetId)
                            {
                                worksetParam.Set(targetWorksetId.IntegerValue);
                                movedCount++;
                            }
                        }
                        LogMessage($"Moved {movedCount} elements to workset '{record.WorksetName}'.");

                        // Store element IDs for this package group for later export
                        string normalizedCode = record.NormalizedPackageCode;
                        if (!packageGroupMapping.ContainsKey(normalizedCode))
                        {
                            packageGroupMapping[normalizedCode] = new List<ElementId>();
                        }
                        packageGroupMapping[normalizedCode].AddRange(elements.Select(e => e.Id));
                    }

                    // 4. HANDLE ORPHANED ELEMENTS (Move to DX_QC)
                    FilteredElementCollector allElementsCollector = new FilteredElementCollector(_doc);
                    var allElementIds = allElementsCollector
                        .WhereElementIsNotElementType()
                        .ToElementIds()
                        .ToHashSet();

                    // Subtract all elements we just processed from the total set
                    var processedElementIds = packageGroupMapping.Values.SelectMany(idList => idList).ToHashSet();
                    allElementIds.ExceptWith(processedElementIds);

                    int orphanedCount = 0;
                    foreach (ElementId orphanId in allElementIds)
                    {
                        Element element = _doc.GetElement(orphanId);
                        RevitParameter worksetParam = element.get_Parameter(BuiltInParameter.ELEM_PARTITION_PARAM);
                        if (worksetParam != null && !worksetParam.IsReadOnly && element.WorksetId != qcWorksetId)
                        {
                            worksetParam.Set(qcWorksetId.IntegerValue);
                            orphanedCount++;
                        }
                    }
                    LogMessage($"Moved {orphanedCount} orphaned elements to 'DX_QC' workset.");

                    trans.Commit();

                    // 5. EXPORT LOGIC
                    if (exportQc)
                    {
                        // Add QC workset to the export groups if user chose to export it
                        var qcElements = new FilteredElementCollector(_doc)
                            .WhereElementIsNotElementType()
                            .Where(e => e.WorksetId == qcWorksetId)
                            .Select(e => e.Id)
                            .ToList();

                        if (qcElements.Any())
                        {
                            packageGroupMapping["QC"] = qcElements;
                        }
                    }

                    // Perform the exports
                    ExportRVTs(packageGroupMapping, destinationPath, overwriteFiles);
                }

                // Write final log file
                System.IO.File.WriteAllText(
                    System.IO.Path.Combine(destinationPath, "WorksetOrchestrationLog.txt"),
                    _log.ToString()
                );

                return true;
            }
            catch (Exception ex)
            {
                LogMessage($"ERROR: {ex.Message}");
                LogMessage($"Stack Trace: {ex.StackTrace}");
                return false;
            }
        }

        private WorksetId GetOrCreateWorkset(string worksetName)
        {
            Workset workset = new FilteredWorksetCollector(_doc)
                .OfKind(WorksetKind.UserWorkset)
                .FirstOrDefault(w => w.Name.Equals(worksetName, StringComparison.CurrentCultureIgnoreCase));

            if (workset != null)
                return workset.Id;

            // Workset doesn't exist, create it
            try
            {
                Workset newWorkset = Workset.Create(_doc, worksetName);
                LogMessage($"Created new workset: {worksetName}");
                return newWorkset.Id;
            }
            catch (Exception ex)
            {
                LogMessage($"ERROR creating workset '{worksetName}': {ex.Message}");
                return WorksetId.InvalidWorksetId;
            }
        }

        private void ExportRVTs(Dictionary<string, List<ElementId>> packageGroups, string destinationPath, bool overwrite)
        {
            string projectPrefix = System.IO.Path.GetFileNameWithoutExtension(_doc.PathName).Split('_')[0];

            foreach (var group in packageGroups)
            {
                string packageCode = group.Key;
                if (packageCode == "NO EXPORT")
                    continue;

                LogMessage($"Preparing export for package: {packageCode}");

                // Create a new temporary Revit file for this export
                string exportFileName = $"{projectPrefix}_{packageCode}_MO_Part_001_DX.rvt";
                string exportFilePath = System.IO.Path.Combine(destinationPath, exportFileName);

                if (System.IO.File.Exists(exportFilePath) && !overwrite)
                {
                    LogMessage($"Skipped - File exists and overwrite is false: {exportFileName}");
                    continue;
                }

                // Use SaveAs with option to create new file
                SaveAsOptions options = new SaveAsOptions();
                options.OverwriteExistingFile = overwrite;

                try
                {
                    // This is a simplified approach - in production, you'd need to
                    // create a new document and selectively copy elements
                    _doc.SaveAs(exportFilePath, options);
                    LogMessage($"Exported: {exportFileName}");
                }
                catch (Exception ex)
                {
                    LogMessage($"ERROR exporting {exportFileName}: {ex.Message}");
                }
            }
        }
    }
}