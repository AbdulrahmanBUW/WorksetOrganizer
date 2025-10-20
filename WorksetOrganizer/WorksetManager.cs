using System;
using System.Linq;
using Autodesk.Revit.DB;

namespace WorksetOrchestrator
{
    public class WorksetManager
    {
        private readonly Document _doc;
        private readonly Action<string> _logAction;

        public WorksetManager(Document doc, Action<string> logAction)
        {
            _doc = doc;
            _logAction = logAction;
        }

        public WorksetId GetOrCreateWorkset(string worksetName)
        {
            Workset workset = new FilteredWorksetCollector(_doc)
                .OfKind(WorksetKind.UserWorkset)
                .FirstOrDefault(w => w.Name.Equals(worksetName, StringComparison.CurrentCultureIgnoreCase));

            if (workset != null)
                return workset.Id;

            try
            {
                var newWorkset = Workset.Create(_doc, worksetName);
                _logAction($"Created new workset: {worksetName}");
                return newWorkset.Id;
            }
            catch (Exception ex)
            {
                _logAction($"ERROR creating workset '{worksetName}': {ex.Message}");
                return WorksetId.InvalidWorksetId;
            }
        }

        public WorksetId GetOrCreateWorksetInDocument(Document doc, string worksetName)
        {
            try
            {
                Workset workset = new FilteredWorksetCollector(doc)
                    .OfKind(WorksetKind.UserWorkset)
                    .FirstOrDefault(w => w.Name.Equals(worksetName, StringComparison.CurrentCultureIgnoreCase));

                if (workset != null)
                    return workset.Id;

                var newWorkset = Workset.Create(doc, worksetName);
                _logAction($"Created new workset in target document: {worksetName}");
                return newWorkset.Id;
            }
            catch (Exception ex)
            {
                _logAction($"ERROR creating workset '{worksetName}' in target document: {ex.Message}");
                return WorksetId.InvalidWorksetId;
            }
        }

        public void AssignElementsToWorkset(Document targetDoc, System.Collections.Generic.List<ElementId> copiedElementIds, string worksetName)
        {
            if (copiedElementIds == null || !copiedElementIds.Any())
                return;

            try
            {
                _logAction($"Assigning {copiedElementIds.Count} elements to workset '{worksetName}' in template");

                if (!targetDoc.IsWorkshared)
                {
                    _logAction($"Template document is not workshared - elements will remain in default workset");
                    return;
                }

                WorksetId targetWorksetId = GetOrCreateWorksetInDocument(targetDoc, worksetName);

                if (targetWorksetId != WorksetId.InvalidWorksetId)
                {
                    int assignedCount = 0;
                    int genericModelCount = 0;

                    foreach (var elementId in copiedElementIds)
                    {
                        try
                        {
                            Element element = targetDoc.GetElement(elementId);
                            if (element != null)
                            {
                                if (element.Category?.Id.IntegerValue == (int)BuiltInCategory.OST_GenericModel)
                                {
                                    genericModelCount++;
                                    _logAction($"  Assigning Generic Model {element.Id} to workset '{worksetName}'");
                                }

                                var worksetParam = element.get_Parameter(BuiltInParameter.ELEM_PARTITION_PARAM);
                                if (worksetParam != null && !worksetParam.IsReadOnly)
                                {
                                    worksetParam.Set(targetWorksetId.IntegerValue);
                                    assignedCount++;
                                }
                                else
                                {
                                    _logAction($"  Warning: Cannot assign element {element.Id} to workset (parameter readonly or missing)");
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            _logAction($"Warning: Could not assign element {elementId} to workset: {ex.Message}");
                        }
                    }

                    _logAction($"Successfully assigned {assignedCount} elements (including {genericModelCount} Generic Models) to workset '{worksetName}' in template");
                }
                else
                {
                    _logAction($"Could not create/find workset '{worksetName}' in template document");
                }
            }
            catch (Exception ex)
            {
                _logAction($"ERROR assigning elements to workset '{worksetName}' in template: {ex.Message}");
            }
        }

        public void SynchronizeWithCentral()
        {
            if (_doc.IsWorkshared)
            {
                try
                {
                    _logAction("Attempting to synchronize with central and relinquish ownership...");

                    var transOpts = new TransactWithCentralOptions();
                    transOpts.SetLockCallback(new SynchLockCallback());

                    var syncOpts = new SynchronizeWithCentralOptions();
                    var relinquishOpts = new RelinquishOptions(true);
                    syncOpts.SetRelinquishOptions(relinquishOpts);
                    syncOpts.SaveLocalAfter = false;

                    _doc.SynchronizeWithCentral(transOpts, syncOpts);

                    _logAction("Synchronized with central and relinquished ownership (per options).");
                }
                catch (Exception ex)
                {
                    _logAction($"WARNING: Synchronization with central failed: {ex.Message}");
                }
            }
            else
            {
                _logAction("Document is not workshared - skipping synchronize/relinquish step.");
            }
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