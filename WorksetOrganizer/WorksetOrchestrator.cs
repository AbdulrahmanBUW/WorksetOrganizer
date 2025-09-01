using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

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
                if (!_doc.IsWorkshared)
                {
                    LogMessage("ERROR: Document is not workshared. Please enable worksharing.");
                    return false;
                }

                var packageGroupMapping = new Dictionary<string, List<ElementId>>();

                using (Transaction trans = new Transaction(_doc, "Process Workset Mapping"))
                {
                    trans.Start();

                    // 1️⃣ Create DX_QC workset inside transaction
                    WorksetId qcWorksetId = GetOrCreateWorkset("DX_QC");

                    foreach (var record in mapping)
                    {
                        LogMessage($"Processing record for system: {record.SystemNameInModel}");

                        if (record.ModelPackageCode == "NO EXPORT")
                        {
                            LogMessage($"Skipped - Marked as 'NO EXPORT'.");
                            continue;
                        }

                        // Create/get target workset inside transaction
                        WorksetId targetWorksetId = GetOrCreateWorkset(record.WorksetName);

                        var collector = new FilteredElementCollector(_doc)
                            .OfClass(typeof(FamilyInstance))
                            .WhereElementIsNotElementType()
                            .Where(e =>
                                e.get_Parameter(BuiltInParameter.RBS_SYSTEM_NAME_PARAM) != null &&
                                e.get_Parameter(BuiltInParameter.RBS_SYSTEM_NAME_PARAM).AsString() == record.SystemNameInModel
                            )
                            .ToList();

                        LogMessage($"Found {collector.Count} elements for system '{record.SystemNameInModel}'.");

                        int movedCount = 0;
                        foreach (Element element in collector)
                        {
                            var worksetParam = element.get_Parameter(BuiltInParameter.ELEM_PARTITION_PARAM);
                            if (worksetParam != null && !worksetParam.IsReadOnly && element.WorksetId != targetWorksetId)
                            {
                                worksetParam.Set(targetWorksetId.IntegerValue);
                                movedCount++;
                            }
                        }

                        LogMessage($"Moved {movedCount} elements to workset '{record.WorksetName}'.");

                        string normalizedCode = record.NormalizedPackageCode;
                        if (!packageGroupMapping.ContainsKey(normalizedCode))
                            packageGroupMapping[normalizedCode] = new List<ElementId>();

                        packageGroupMapping[normalizedCode].AddRange(collector.Select(e => e.Id));
                    }

                    // Move orphaned elements to DX_QC
                    var allElementIds = new FilteredElementCollector(_doc)
                        .WhereElementIsNotElementType()
                        .ToElementIds()
                        .ToHashSet();

                    var processedElementIds = packageGroupMapping.Values.SelectMany(x => x).ToHashSet();
                    allElementIds.ExceptWith(processedElementIds);

                    int orphanedCount = 0;
                    foreach (var orphanId in allElementIds)
                    {
                        var element = _doc.GetElement(orphanId);
                        var worksetParam = element.get_Parameter(BuiltInParameter.ELEM_PARTITION_PARAM);
                        if (worksetParam != null && !worksetParam.IsReadOnly && element.WorksetId != qcWorksetId)
                        {
                            worksetParam.Set(qcWorksetId.IntegerValue);
                            orphanedCount++;
                        }
                    }

                    LogMessage($"Moved {orphanedCount} orphaned elements to 'DX_QC' workset.");

                    trans.Commit();
                }

                // Export logic outside transaction
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
            string projectPrefix = System.IO.Path.GetFileNameWithoutExtension(_doc.PathName).Split('_')[0];

            if (exportQc)
            {
                var qcElements = new FilteredElementCollector(_doc)
                    .WhereElementIsNotElementType()
                    .Where(e => e.WorksetId == GetOrCreateWorkset("DX_QC"))
                    .Select(e => e.Id)
                    .ToList();

                if (qcElements.Any())
                    packageGroups["QC"] = qcElements;
            }

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
