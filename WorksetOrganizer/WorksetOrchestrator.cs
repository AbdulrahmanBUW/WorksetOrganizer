using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace WorksetOrchestrator
{
    public class WorksetOrchestrator
    {
        private Document _doc;
        private UIDocument _uiDoc;
        private StringBuilder _log = new StringBuilder();

        private readonly WorksetMapper _worksetMapper;
        private readonly ElementClassifier _classifier;
        private readonly ElementMatcher _matcher;
        private readonly ElementCollectorService _collectorService;
        private readonly WorksetManager _worksetManager;
        private readonly DocumentExportService _exportService;
        private readonly TemplateIntegrationService _templateService;

        public event EventHandler<string> LogUpdated;

        public WorksetOrchestrator(UIDocument uiDoc)
        {
            _uiDoc = uiDoc;
            _doc = uiDoc.Document;

            _worksetMapper = new WorksetMapper();
            _classifier = new ElementClassifier(LogMessage);
            _matcher = new ElementMatcher(LogMessage, _classifier);
            _collectorService = new ElementCollectorService(_doc, LogMessage);
            _worksetManager = new WorksetManager(_doc, LogMessage);
            _exportService = new DocumentExportService(_doc, _uiDoc, LogMessage, _worksetMapper);
            _templateService = new TemplateIntegrationService(_uiDoc, LogMessage, _worksetMapper);
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

                EnsureDefaultMappings(mapping);

                using (Transaction trans = new Transaction(_doc, "Process Workset Mapping"))
                {
                    trans.Start();

                    WorksetId qcWorksetId = _worksetManager.GetOrCreateWorkset("DX_QC");

                    LogMessage("=== STEP 1: Preserving existing workset assignments ===");
                    var existingWorksetElements = _collectorService.CollectExistingWorksetElements(mapping, packageGroupMapping, _worksetMapper);

                    LogMessage("=== STEP 2: Processing Excel mapping for unassigned elements ===");
                    ProcessMappingRecords(mapping, existingWorksetElements, packageGroupMapping);

                    LogMessage("=== STEP 3: Moving orphaned elements to DX_QC ===");
                    MoveOrphanedElementsToQC(existingWorksetElements, packageGroupMapping, qcWorksetId);

                    trans.Commit();
                }

                _worksetManager.SynchronizeWithCentral();
                _exportService.ExportRVTs(packageGroupMapping, destinationPath, overwriteFiles, exportQc);

                File.WriteAllText(Path.Combine(destinationPath, "WorksetOrchestrationLog.txt"), _log.ToString());

                return true;
            }
            catch (Exception ex)
            {
                LogMessage($"ERROR: {ex.Message}");
                LogMessage($"Stack Trace: {ex.StackTrace}");
                return false;
            }
        }

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

                var allWorksets = new FilteredWorksetCollector(_doc)
                    .OfKind(WorksetKind.UserWorkset)
                    .Where(w => w.Kind == WorksetKind.UserWorkset)
                    .ToList();

                LogMessage($"Found {allWorksets.Count} user worksets in the model:");
                foreach (var workset in allWorksets)
                {
                    LogMessage($"  - {workset.Name} (ID: {workset.Id})");
                }

                var worksetElementMapping = new Dictionary<string, List<ElementId>>();

                foreach (var workset in allWorksets)
                {
                    LogMessage($"Collecting elements for workset: {workset.Name}");

                    var elementsInWorkset = new FilteredElementCollector(_doc)
                        .WhereElementIsNotElementType()
                        .Where(e => e.WorksetId == workset.Id)
                        .ToList();

                    var relevantElements = _collectorService.GetRelevantElementsFromCollection(elementsInWorkset);

                    LogMessage($"  Found {elementsInWorkset.Count} total elements, {relevantElements.Count} relevant elements");

                    if (relevantElements.Count > 0)
                    {
                        worksetElementMapping[workset.Name] = relevantElements.Select(e => e.Id).ToList();

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

                _worksetManager.SynchronizeWithCentral();
                _exportService.ExportWorksetRVTs(worksetElementMapping, destinationPath, overwriteFiles);

                File.WriteAllText(Path.Combine(destinationPath, "WorksetExtractionLog.txt"), _log.ToString());

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

        public bool IntegrateIntoTemplate(List<string> extractedFiles, string templateFilePath, string destinationPath)
        {
            return _templateService.IntegrateIntoTemplate(extractedFiles, templateFilePath, destinationPath);
        }

        private void EnsureDefaultMappings(List<MappingRecord> mapping)
        {
            var eltRecord = mapping.FirstOrDefault(m => m.WorksetName?.ToUpper() == "DX_ELT");
            if (eltRecord == null)
            {
                mapping.Add(new MappingRecord
                {
                    WorksetName = "DX_ELT",
                    SystemNameInModel = "ELT",
                    SystemDescription = "Electrical Elements",
                    ModelPackageCode = "ELT"
                });
                LogMessage("Added default DX_ELT mapping record");
            }
        }

        private void ProcessMappingRecords(List<MappingRecord> mapping, HashSet<ElementId> existingWorksetElements, Dictionary<string, List<ElementId>> packageGroupMapping)
        {
            foreach (var record in mapping)
            {
                LogMessage($"Processing record for system pattern: {record.SystemNameInModel}");

                if (record.ModelPackageCode == "NO EXPORT")
                {
                    LogMessage($"Skipped - Marked as 'NO EXPORT'.");
                    continue;
                }

                WorksetId targetWorksetId = _worksetManager.GetOrCreateWorkset(record.WorksetName);

                var allMepElements = _collectorService.GetAllMepElements();
                var unassignedElements = allMepElements.Where(e => !existingWorksetElements.Contains(e.Id)).ToList();

                LogMessage($"Found {unassignedElements.Count} unassigned elements to check for pattern '{record.SystemNameInModel}'");

                var matchingElements = _matcher.FindMatchingElements(unassignedElements, record);

                LogMessage($"Found {matchingElements.Count} elements for system pattern '{record.SystemNameInModel}'.");

                int movedCount = 0;
                foreach (Element element in matchingElements)
                {
                    var worksetParam = element.get_Parameter(BuiltInParameter.ELEM_PARTITION_PARAM);
                    if (worksetParam != null && !worksetParam.IsReadOnly && element.WorksetId != targetWorksetId)
                    {
                        worksetParam.Set(targetWorksetId.IntegerValue);
                        movedCount++;

                        string normalizedCode = record.NormalizedPackageCode;
                        if (!packageGroupMapping.ContainsKey(normalizedCode))
                            packageGroupMapping[normalizedCode] = new List<ElementId>();
                        packageGroupMapping[normalizedCode].Add(element.Id);
                    }
                }

                LogMessage($"Moved {movedCount} elements to workset '{record.WorksetName}'.");
            }
        }

        private void MoveOrphanedElementsToQC(HashSet<ElementId> existingWorksetElements, Dictionary<string, List<ElementId>> packageGroupMapping, WorksetId qcWorksetId)
        {
            var allMepElementsAfterProcessing = _collectorService.GetAllMepElements();
            var allProcessedElementIds = existingWorksetElements.Concat(
                packageGroupMapping.Values.SelectMany(x => x)).ToHashSet();

            int orphanedCount = 0;
            foreach (var element in allMepElementsAfterProcessing)
            {
                if (!allProcessedElementIds.Contains(element.Id))
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
        }
    }
}