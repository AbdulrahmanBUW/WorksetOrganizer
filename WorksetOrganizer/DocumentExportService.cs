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

        // Categories that typically contain non-copyable or problematic elements
        private static readonly HashSet<BuiltInCategory> ProblematicCategories = new HashSet<BuiltInCategory>
        {
            BuiltInCategory.OST_RvtLinks,
            BuiltInCategory.OST_Views,
            BuiltInCategory.OST_Sheets,
            BuiltInCategory.OST_Schedules,
            BuiltInCategory.OST_ProjectInformation,
            BuiltInCategory.OST_SharedBasePoint,
            BuiltInCategory.OST_ProjectBasePoint,
            BuiltInCategory.OST_Cameras,
            BuiltInCategory.OST_Matchline,
            BuiltInCategory.OST_VolumeOfInterest,  // This is Scope Boxes in Revit API
            BuiltInCategory.OST_Viewers,
            BuiltInCategory.OST_Grids,
            BuiltInCategory.OST_Levels,
            BuiltInCategory.OST_CLines,  // Reference Planes in Revit API
            BuiltInCategory.OST_ReferenceLines,
            BuiltInCategory.OST_AreaSchemes,
            BuiltInCategory.OST_Areas,
            BuiltInCategory.OST_Rooms,
            BuiltInCategory.OST_RoomSeparationLines,
            BuiltInCategory.OST_MEPSpaceSeparationLines,
            BuiltInCategory.OST_CurtainGrids,
            BuiltInCategory.OST_CurtainGridsRoof,
            BuiltInCategory.OST_CurtainGridsSystem,
            BuiltInCategory.OST_CurtainGridsWall
        };

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

            // Sort worksets alphabetically for consistent part numbering
            var sortedWorksets = worksetMapping.OrderBy(w => w.Key, StringComparer.OrdinalIgnoreCase).ToList();

            // Track iFLS codes for part numbering
            var iflsCodeCounts = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

            _logAction($"Preparing to export {sortedWorksets.Count} workset files...");

            int exportedCount = 0;
            var app = _uiDoc.Application.Application;

            foreach (var worksetGroup in sortedWorksets)
            {
                string worksetName = worksetGroup.Key;
                var elementIds = worksetGroup.Value;

                if (elementIds == null || !elementIds.Any())
                {
                    _logAction($"Skipping '{worksetName}' - no elements to export");
                    continue;
                }

                string iflsCode = _worksetMapper.GetIflsCodeForWorkset(worksetName);

                // Calculate part number for this iFLS code
                if (!iflsCodeCounts.ContainsKey(iflsCode))
                {
                    iflsCodeCounts[iflsCode] = 0;
                }
                iflsCodeCounts[iflsCode]++;
                int partNumber = iflsCodeCounts[iflsCode];

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

        /// <summary>
        /// Filters out elements that cannot be copied to another document.
        /// </summary>
        private List<ElementId> FilterCopyableElements(List<ElementId> elementIds, string worksetName)
        {
            var copyableIds = new List<ElementId>();
            int skippedNull = 0;
            int skippedInvalid = 0;
            int skippedNoCategory = 0;
            int skippedProblematic = 0;
            int skippedElementType = 0;
            int skippedViewSpecific = 0;
            int skippedOther = 0;

            foreach (var id in elementIds)
            {
                if (id == null)
                {
                    skippedNull++;
                    continue;
                }

                if (id == ElementId.InvalidElementId)
                {
                    skippedInvalid++;
                    continue;
                }

                try
                {
                    Element element = _doc.GetElement(id);
                    if (element == null)
                    {
                        skippedNull++;
                        continue;
                    }

                    // Skip element types - they are copied automatically with instances
                    if (element is ElementType)
                    {
                        skippedElementType++;
                        continue;
                    }

                    // Skip views and view-specific elements
                    if (element is View || element.OwnerViewId != ElementId.InvalidElementId)
                    {
                        skippedViewSpecific++;
                        continue;
                    }

                    // Skip elements without category (these typically can't be copied)
                    if (element.Category == null)
                    {
                        skippedNoCategory++;
                        continue;
                    }

                    // Skip problematic categories
                    var categoryId = (BuiltInCategory)element.Category.Id.IntegerValue;
                    if (ProblematicCategories.Contains(categoryId))
                    {
                        skippedProblematic++;
                        continue;
                    }

                    // Additional checks for specific element types that may cause issues
                    if (!IsElementCopyable(element))
                    {
                        skippedOther++;
                        continue;
                    }

                    copyableIds.Add(id);
                }
                catch (Exception ex)
                {
                    _logAction($"    Warning: Error checking element {id}: {ex.Message}");
                    skippedOther++;
                }
            }

            // Log summary of filtered elements
            int totalSkipped = skippedNull + skippedInvalid + skippedNoCategory + skippedProblematic +
                               skippedElementType + skippedViewSpecific + skippedOther;

            if (totalSkipped > 0)
            {
                _logAction($"  Filtered out {totalSkipped} non-copyable elements for '{worksetName}':");
                if (skippedNull > 0) _logAction($"    - Null/missing elements: {skippedNull}");
                if (skippedInvalid > 0) _logAction($"    - Invalid IDs: {skippedInvalid}");
                if (skippedNoCategory > 0) _logAction($"    - No category: {skippedNoCategory}");
                if (skippedProblematic > 0) _logAction($"    - Problematic categories: {skippedProblematic}");
                if (skippedElementType > 0) _logAction($"    - Element types: {skippedElementType}");
                if (skippedViewSpecific > 0) _logAction($"    - View-specific: {skippedViewSpecific}");
                if (skippedOther > 0) _logAction($"    - Other reasons: {skippedOther}");
            }

            return copyableIds;
        }

        /// <summary>
        /// Performs additional checks to determine if an element can be safely copied.
        /// </summary>
        private bool IsElementCopyable(Element element)
        {
            try
            {
                // Check for design options - elements in non-primary design options can be problematic
                var designOptionParam = element.get_Parameter(BuiltInParameter.DESIGN_OPTION_ID);
                if (designOptionParam != null)
                {
                    var designOptionId = designOptionParam.AsElementId();
                    if (designOptionId != null && designOptionId != ElementId.InvalidElementId)
                    {
                        // Element is in a design option - could be problematic
                        // For now, we'll allow it but this could be made more restrictive
                    }
                }

                // Check for group membership - grouped elements should be copied via their group
                var groupId = element.GroupId;
                if (groupId != null && groupId != ElementId.InvalidElementId)
                {
                    // Element is in a group - we'll still try to copy it
                    // The group handling in Revit should manage this
                }

                // Check if element is a system family that requires host
                if (element is FamilyInstance familyInstance)
                {
                    // Check if it's a hosted element without a valid host
                    var host = familyInstance.Host;
                    if (host != null)
                    {
                        // Element has a host - it should be copyable if the host is also being copied
                        // This is usually handled automatically by CopyElements
                    }
                }

                // Check for MEP system elements
                if (element is MEPSystem)
                {
                    // MEP Systems should typically be copied via their components
                    return false;
                }

                return true;
            }
            catch
            {
                // If we can't determine, assume it's copyable and let CopyElements handle it
                return true;
            }
        }

        /// <summary>
        /// Copies elements in chunks to isolate problematic elements.
        /// Falls back to individual element copying if chunks fail.
        /// </summary>
        private ICollection<ElementId> CopyElementsInChunks(List<ElementId> elementIds, Document targetDoc,
            CopyPasteOptions copyOptions, string worksetName)
        {
            var allCopiedIds = new List<ElementId>();
            const int chunkSize = 50; // Process 50 elements at a time

            var chunks = elementIds
                .Select((id, index) => new { id, index })
                .GroupBy(x => x.index / chunkSize)
                .Select(g => g.Select(x => x.id).ToList())
                .ToList();

            _logAction($"  Splitting {elementIds.Count} elements into {chunks.Count} chunks of up to {chunkSize} elements");

            int chunkNumber = 0;
            foreach (var chunk in chunks)
            {
                chunkNumber++;
                try
                {
                    var copiedChunk = ElementTransformUtils.CopyElements(
                        _doc,
                        chunk,
                        targetDoc,
                        Transform.Identity,
                        copyOptions
                    );

                    if (copiedChunk != null && copiedChunk.Any())
                    {
                        allCopiedIds.AddRange(copiedChunk);
                        _logAction($"    Chunk {chunkNumber}/{chunks.Count}: Copied {copiedChunk.Count} elements");
                    }
                }
                catch (Exception chunkEx)
                {
                    _logAction($"    Chunk {chunkNumber}/{chunks.Count} failed: {chunkEx.Message}");
                    _logAction($"    Attempting individual element copying for this chunk...");

                    // Fall back to individual element copying for this chunk
                    var individualCopied = CopyElementsIndividually(chunk, targetDoc, copyOptions, worksetName);
                    allCopiedIds.AddRange(individualCopied);
                }
            }

            return allCopiedIds;
        }

        /// <summary>
        /// Copies elements one by one, skipping any that fail.
        /// This is the ultimate fallback to ensure maximum element recovery.
        /// </summary>
        private List<ElementId> CopyElementsIndividually(List<ElementId> elementIds, Document targetDoc,
            CopyPasteOptions copyOptions, string worksetName)
        {
            var copiedIds = new List<ElementId>();
            int successCount = 0;
            int failCount = 0;
            var failedCategories = new Dictionary<string, int>();

            foreach (var elementId in elementIds)
            {
                try
                {
                    var singleElementList = new List<ElementId> { elementId };
                    var copied = ElementTransformUtils.CopyElements(
                        _doc,
                        singleElementList,
                        targetDoc,
                        Transform.Identity,
                        copyOptions
                    );

                    if (copied != null && copied.Any())
                    {
                        copiedIds.AddRange(copied);
                        successCount++;
                    }
                }
                catch (Exception)
                {
                    failCount++;

                    // Track which categories are failing
                    try
                    {
                        var element = _doc.GetElement(elementId);
                        string categoryName = element?.Category?.Name ?? "Unknown";
                        if (!failedCategories.ContainsKey(categoryName))
                            failedCategories[categoryName] = 0;
                        failedCategories[categoryName]++;
                    }
                    catch { }
                }
            }

            _logAction($"      Individual copy results: {successCount} succeeded, {failCount} failed");

            if (failedCategories.Any())
            {
                _logAction($"      Failed elements by category:");
                foreach (var kvp in failedCategories.OrderByDescending(x => x.Value).Take(5))
                {
                    _logAction($"        - {kvp.Key}: {kvp.Value}");
                }
            }

            return copiedIds;
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

                // Filter copyable elements
                var sourceElementIds = FilterCopyableElements(group.Value.ToList(), group.Key);

                if (!sourceElementIds.Any())
                {
                    _logAction($"No copyable elements for group '{group.Key}' - skipping.");
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

                        ICollection<ElementId> copiedIds = null;

                        // Try batch copy first, fall back to chunked if needed
                        try
                        {
                            copiedIds = ElementTransformUtils.CopyElements(
                                _doc,
                                sourceElementIds,
                                newDoc,
                                Transform.Identity,
                                copyOptions
                            );
                        }
                        catch (Exception batchEx)
                        {
                            _logAction($"Batch copy failed for '{group.Key}': {batchEx.Message}");
                            _logAction($"Falling back to chunked copying...");
                            copiedIds = CopyElementsInChunks(sourceElementIds, newDoc, copyOptions, group.Key);
                        }

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

                // Validate and filter element IDs - remove non-copyable elements
                var copyableElementIds = FilterCopyableElements(elementIds, worksetName);

                _logAction($"Created temporary document for workset '{worksetName}' - {copyableElementIds.Count} copyable elements (from {elementIds.Count} total)");

                if (!copyableElementIds.Any())
                {
                    _logAction($"No copyable element IDs for workset '{worksetName}' - skipping.");
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

                        // First attempt: Try to copy all elements at once
                        try
                        {
                            copiedIds = ElementTransformUtils.CopyElements(
                                _doc,
                                copyableElementIds,
                                newDoc,
                                Transform.Identity,
                                copyOptions
                            );
                            _logAction($"Batch copy successful: {copiedIds?.Count ?? 0} elements copied");
                        }
                        catch (Exception batchEx)
                        {
                            _logAction($"Batch copy failed: {batchEx.Message}");
                            _logAction($"Falling back to chunked copying strategy...");

                            // Fallback: Copy in smaller chunks to isolate problematic elements
                            copiedIds = CopyElementsInChunks(copyableElementIds, newDoc, copyOptions, worksetName);
                        }

                        tNew.Commit();

                        _logAction($"Total copied: {copiedIds?.Count ?? 0} elements for workset '{worksetName}'");
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