using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace WorksetOrchestrator
{
    public class DocumentExportService
    {
        private readonly Document _doc;
        private readonly UIDocument _uiDoc;
        private readonly Action<string> _logAction;
        private readonly WorksetMapper _worksetMapper;

        public DocumentExportService(Document doc, UIDocument uiDoc, Action<string> logAction, WorksetMapper worksetMapper)
        {
            _doc = doc;
            _uiDoc = uiDoc;
            _logAction = logAction;
            _worksetMapper = worksetMapper;
        }

        public void ExportRVTs(Dictionary<string, List<ElementId>> packageGroups, string destinationPath, bool overwrite, bool exportQc)
        {
            string projectPrefix = Path.GetFileNameWithoutExtension(_doc.PathName);
            if (string.IsNullOrEmpty(projectPrefix))
            {
                projectPrefix = _doc.Title;
            }

            if (projectPrefix.Contains('_'))
                projectPrefix = projectPrefix.Split('_')[0];

            if (exportQc)
            {
                AddQcElementsToExport(packageGroups);
            }

            RemoveEmptyGroups(packageGroups);

            _logAction($"Preparing to export {packageGroups.Count} model files...");

            int exportedCount = 0;
            foreach (var group in packageGroups)
            {
                if (string.Equals(group.Key, "NO EXPORT", StringComparison.OrdinalIgnoreCase))
                    continue;

                string iflsCode = _worksetMapper.GetIflsCodeForPackage(group.Key);
                string exportFileName = $"{projectPrefix}_{iflsCode}_MO_Part_001_DX.rvt";
                string exportFilePath = Path.Combine(destinationPath, exportFileName);

                _logAction($"Exporting package '{group.Key}' -> iFLS code '{iflsCode}' -> file '{exportFileName}'");

                if (File.Exists(exportFilePath) && !overwrite)
                {
                    _logAction($"Skipped - File exists and overwrite is false: {exportFileName}");
                    continue;
                }

                if (ExportPackageToFile(group, exportFilePath, overwrite))
                {
                    exportedCount++;
                }
            }

            if (exportedCount == 0)
            {
                _logAction("WARNING: No model files were exported - no matching elements found or all exports skipped.");
            }
            else
            {
                _logAction($"Completed exports: {exportedCount} files created.");
            }
        }

        public void ExportWorksetRVTs(Dictionary<string, List<ElementId>> worksetMapping, string destinationPath, bool overwrite)
        {
            string projectPrefix = Path.GetFileNameWithoutExtension(_doc.PathName);
            if (string.IsNullOrEmpty(projectPrefix))
            {
                projectPrefix = _doc.Title;
            }

            if (projectPrefix.Contains('_'))
                projectPrefix = projectPrefix.Split('_')[0];

            _logAction($"Preparing to export {worksetMapping.Count} workset files...");

            int exportedCount = 0;
            var app = _uiDoc.Application.Application;

            foreach (var worksetGroup in worksetMapping)
            {
                string worksetName = worksetGroup.Key;
                var elementIds = worksetGroup.Value;

                if (elementIds == null || !elementIds.Any())
                {
                    _logAction($"Skipping '{worksetName}' - no elements to export");
                    continue;
                }

                string iflsCode = _worksetMapper.GetIflsCodeForWorkset(worksetName);

                // HARDCODED: PWI = Part_001, UPW = Part_002
                int partNumber = 1; // Default
                if (worksetName.Equals("DX_PWI", StringComparison.OrdinalIgnoreCase))
                {
                    partNumber = 1;
                }
                else if (worksetName.Equals("DX_UPW", StringComparison.OrdinalIgnoreCase))
                {
                    partNumber = 2;
                }

                string exportFileName = $"{projectPrefix}_{iflsCode}_MO_Part_{partNumber:D3}_DX.rvt";
                string exportFilePath = Path.Combine(destinationPath, exportFileName);

                _logAction($"Processing workset '{worksetName}' -> iFLS code '{iflsCode}' -> Part {partNumber:D3} -> file '{exportFileName}'");

                if (File.Exists(exportFilePath) && !overwrite)
                {
                    _logAction($"Skipped - File exists: {exportFileName}");
                    continue;
                }

                if (ExportWorksetToFile(worksetName, elementIds, exportFilePath, overwrite, app))
                {
                    exportedCount++;
                }
            }

            if (exportedCount == 0)
            {
                _logAction("WARNING: No workset files were exported - no relevant elements found or all exports skipped.");
            }
            else
            {
                _logAction($"Workset extraction completed: {exportedCount} files created with correct iFLS naming.");
            }
        }

        private void AddQcElementsToExport(Dictionary<string, List<ElementId>> packageGroups)
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
                    _logAction($"Added {qcElements.Count} elements to QC export group.");
                }
                else
                {
                    _logAction("No elements found in DX_QC workset - skipping QC export.");
                }
            }
        }

        private void RemoveEmptyGroups(Dictionary<string, List<ElementId>> packageGroups)
        {
            var emptyGroups = packageGroups.Where(g => g.Value == null || !g.Value.Any()).Select(g => g.Key).ToList();
            foreach (var emptyKey in emptyGroups)
            {
                packageGroups.Remove(emptyKey);
                _logAction($"Skipping export for '{emptyKey}' - no elements found.");
            }
        }

        private bool ExportPackageToFile(KeyValuePair<string, List<ElementId>> group, string exportFilePath, bool overwrite)
        {
            Document newDoc = null;
            try
            {
                string templatePath = null;
                try
                {
                    templatePath = _uiDoc.Application.Application.DefaultProjectTemplate;
                }
                catch
                {
                    templatePath = null;
                }

                if (string.IsNullOrEmpty(templatePath) || !File.Exists(templatePath))
                {
                    _logAction("Default project template not available or not found - creating new blank project document.");
                    newDoc = _uiDoc.Application.Application.NewProjectDocument(UnitSystem.Metric);
                }
                else
                {
                    newDoc = _uiDoc.Application.Application.NewProjectDocument(templatePath);
                }

                if (newDoc == null)
                    throw new Exception("Failed to create new project document.");

                _logAction($"Created temporary new project document for export '{group.Key}'.");

                var sourceElementIds = group.Value.Where(id => id != null && id != ElementId.InvalidElementId).Distinct().ToList();

                if (!sourceElementIds.Any())
                {
                    _logAction($"No valid elements to copy for group '{group.Key}' - skipping.");
                    try { newDoc.Close(false); } catch { }
                    return false;
                }

                try
                {
                    using (Transaction tNew = new Transaction(newDoc, $"Copy elements for {group.Key}"))
                    {
                        tNew.Start();

                        var copyOptions = new CopyPasteOptions();
                        copyOptions.SetDuplicateTypeNamesHandler(new DuplicateTypeNamesHandler());

                        ICollection<ElementId> copiedIds = ElementTransformUtils.CopyElements(
                            _doc,
                            sourceElementIds,
                            newDoc,
                            Transform.Identity,
                            copyOptions
                        );

                        tNew.Commit();

                        _logAction($"Copied {copiedIds?.Count ?? 0} elements into temporary document for group '{group.Key}'.");
                    }
                }
                catch (Exception copyEx)
                {
                    _logAction($"ERROR copying elements for group '{group.Key}': {copyEx.Message}");
                    try { newDoc.Close(false); } catch { }
                    return false;
                }

                try
                {
                    var saveOpts = new SaveAsOptions
                    {
                        OverwriteExistingFile = overwrite
                    };

                    newDoc.SaveAs(exportFilePath, saveOpts);
                    _logAction($"Exported: {Path.GetFileName(exportFilePath)} with {group.Value.Count} elements");
                    return true;
                }
                catch (Exception saveEx)
                {
                    _logAction($"ERROR saving exported file '{Path.GetFileName(exportFilePath)}': {saveEx.Message}");
                    return false;
                }
            }
            catch (Exception ex)
            {
                _logAction($"ERROR preparing export for group '{group.Key}': {ex.Message}");
                return false;
            }
            finally
            {
                try
                {
                    if (newDoc != null)
                    {
                        bool closeResult = newDoc.Close(false);
                        _logAction($"Temporary document closed for group '{group.Key}' (success: {closeResult}).");
                    }
                }
                catch (Exception closeEx)
                {
                    _logAction($"Warning: could not close temporary document for '{group.Key}': {closeEx.Message}");
                }
            }
        }

        private bool ExportWorksetToFile(string worksetName, List<ElementId> elementIds, string exportFilePath, bool overwrite, Application app)
        {
            Document newDoc = null;
            try
            {
                _logAction($"Creating new document for workset: {worksetName}");

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

                _logAction($"Created temporary document for workset '{worksetName}' with {elementIds.Count} elements");

                // Validate element IDs
                var validElementIds = elementIds.Where(id => id != null && id != ElementId.InvalidElementId).ToList();
                if (!validElementIds.Any())
                {
                    _logAction($"No valid element IDs for workset '{worksetName}' - skipping.");
                    try { newDoc.Close(false); } catch { }
                    return false;
                }

                ICollection<ElementId> copiedIds = null;
                try
                {
                    using (Transaction tNew = new Transaction(newDoc, $"Copy workset {worksetName}"))
                    {
                        tNew.Start();

                        var copyOptions = new CopyPasteOptions();
                        copyOptions.SetDuplicateTypeNamesHandler(new DuplicateTypeNamesHandler());

                        copiedIds = ElementTransformUtils.CopyElements(
                            _doc,
                            validElementIds,
                            newDoc,
                            Transform.Identity,
                            copyOptions
                        );

                        tNew.Commit();

                        _logAction($"Copied {copiedIds?.Count ?? 0} elements for workset '{worksetName}'");
                    }
                }
                catch (Exception copyEx)
                {
                    _logAction($"ERROR copying elements for workset '{worksetName}': {copyEx.Message}");
                    try { newDoc.Close(false); } catch { }
                    return false;
                }

                if (copiedIds != null && copiedIds.Any())
                {
                    var worksetManager = new WorksetManager(newDoc, _logAction);
                    using (Transaction trans = new Transaction(newDoc, $"Assign elements to workset {worksetName}"))
                    {
                        trans.Start();
                        worksetManager.AssignElementsToWorkset(newDoc, copiedIds.ToList(), worksetName);
                        trans.Commit();
                    }
                }

                try
                {
                    var saveOpts = new SaveAsOptions { OverwriteExistingFile = overwrite };

                    if (newDoc.IsWorkshared)
                    {
                        try
                        {
                            var wsOpts = new WorksharingSaveAsOptions { SaveAsCentral = true };
                            saveOpts.SetWorksharingOptions(wsOpts);
                            _logAction($"Saving as central file to preserve worksets");
                        }
                        catch (Exception wsEx)
                        {
                            _logAction($"Warning: Could not set worksharing save options: {wsEx.Message}");
                        }
                    }

                    newDoc.SaveAs(exportFilePath, saveOpts);
                    _logAction($"Exported: {Path.GetFileName(exportFilePath)} (workset: {worksetName} -> iFLS: {_worksetMapper.GetIflsCodeForWorkset(worksetName)})");
                    return true;
                }
                catch (Exception saveEx)
                {
                    _logAction($"ERROR saving workset file '{Path.GetFileName(exportFilePath)}': {saveEx.Message}");
                    return false;
                }
            }
            catch (Exception ex)
            {
                _logAction($"ERROR preparing export for workset '{worksetName}': {ex.Message}");
                return false;
            }
            finally
            {
                try
                {
                    if (newDoc != null)
                    {
                        bool closeResult = newDoc.Close(false);
                        _logAction($"Closed temporary document for workset '{worksetName}' (success: {closeResult})");
                    }
                }
                catch (Exception closeEx)
                {
                    _logAction($"Warning: Could not close temporary document for '{worksetName}': {closeEx.Message}");
                }
            }
        }
    }
}