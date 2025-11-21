using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;

namespace WorksetOrchestrator
{
    public class ElementCollectorService
    {
        private readonly Document _doc;
        private readonly Action<string> _logAction;

        // Categories that should NOT be collected for export
        private static readonly HashSet<BuiltInCategory> ExcludedCategories = new HashSet<BuiltInCategory>
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
            BuiltInCategory.OST_AreaSchemes,
            BuiltInCategory.OST_MEPSpaceSeparationLines,
            BuiltInCategory.OST_CurtainGrids,
            BuiltInCategory.OST_CurtainGridsRoof,
            BuiltInCategory.OST_CurtainGridsSystem,
            BuiltInCategory.OST_CurtainGridsWall
        };

        public ElementCollectorService(Document doc, Action<string> logAction)
        {
            _doc = doc;
            _logAction = logAction;
        }

        public List<Element> GetAllMepElements()
        {
            var mepElements = new List<Element>();

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
                BuiltInCategory.OST_ConduitFitting,
                BuiltInCategory.OST_GenericModel,
                BuiltInCategory.OST_SpecialityEquipment,
                BuiltInCategory.OST_Mass,
                BuiltInCategory.OST_Furniture,
                BuiltInCategory.OST_FurnitureSystems
            };

            foreach (var category in categories)
            {
                try
                {
                    var collector = new FilteredElementCollector(_doc)
                        .OfCategory(category)
                        .WhereElementIsNotElementType();

                    var elementsInCategory = collector.ToList();
                    mepElements.AddRange(elementsInCategory);

                    if (category == BuiltInCategory.OST_GenericModel)
                    {
                        _logAction($"Found {elementsInCategory.Count} Generic Model elements");
                    }
                }
                catch (Exception ex)
                {
                    _logAction($"Warning: Could not collect category {category}: {ex.Message}");
                }
            }

            _logAction($"Collected {mepElements.Count} total elements from {categories.Count} categories.");
            return mepElements;
        }

        public List<Element> GetAllMepElementsFromDocument(Document doc)
        {
            var mepElements = new List<Element>();

            // COMPREHENSIVE list of all possible categories
            var categories = new List<BuiltInCategory>
            {
                // Piping
                BuiltInCategory.OST_PipeFitting,
                BuiltInCategory.OST_PipeAccessory,
                BuiltInCategory.OST_PipeCurves,
                BuiltInCategory.OST_FlexPipeCurves,
                BuiltInCategory.OST_PipeInsulations,
                
                // Ductwork
                BuiltInCategory.OST_DuctFitting,
                BuiltInCategory.OST_DuctAccessory,
                BuiltInCategory.OST_DuctCurves,
                BuiltInCategory.OST_FlexDuctCurves,
                BuiltInCategory.OST_DuctTerminal,
                BuiltInCategory.OST_DuctInsulations,
                
                // Mechanical
                BuiltInCategory.OST_MechanicalEquipment,
                BuiltInCategory.OST_PlumbingFixtures,
                BuiltInCategory.OST_Sprinklers,
                BuiltInCategory.OST_FireAlarmDevices,
                
                // Electrical
                BuiltInCategory.OST_ElectricalEquipment,
                BuiltInCategory.OST_ElectricalFixtures,
                BuiltInCategory.OST_LightingFixtures,
                BuiltInCategory.OST_LightingDevices,
                BuiltInCategory.OST_CableTray,
                BuiltInCategory.OST_CableTrayFitting,
                BuiltInCategory.OST_Conduit,
                BuiltInCategory.OST_ConduitFitting,
                BuiltInCategory.OST_DataDevices,
                BuiltInCategory.OST_CommunicationDevices,
                BuiltInCategory.OST_SecurityDevices,
                BuiltInCategory.OST_NurseCallDevices,
                BuiltInCategory.OST_TelephoneDevices,
                
                // Structural
                BuiltInCategory.OST_StructuralFraming,
                BuiltInCategory.OST_StructuralColumns,
                BuiltInCategory.OST_StructuralFoundation,
                BuiltInCategory.OST_StructuralFramingSystem,
                BuiltInCategory.OST_StructuralStiffener,
                BuiltInCategory.OST_StructuralTruss,
                
                // Architectural
                BuiltInCategory.OST_Walls,
                BuiltInCategory.OST_Doors,
                BuiltInCategory.OST_Windows,
                
                // General
                BuiltInCategory.OST_GenericModel,
                BuiltInCategory.OST_SpecialityEquipment,
                BuiltInCategory.OST_Mass,
                BuiltInCategory.OST_Furniture,
                BuiltInCategory.OST_FurnitureSystems,
                BuiltInCategory.OST_Casework,
                BuiltInCategory.OST_Parts
            };

            foreach (var category in categories)
            {
                try
                {
                    var collector = new FilteredElementCollector(doc)
                        .OfCategory(category)
                        .WhereElementIsNotElementType();

                    var elementsInCategory = collector.ToList();
                    mepElements.AddRange(elementsInCategory);

                    if (category == BuiltInCategory.OST_GenericModel && elementsInCategory.Count > 0)
                    {
                        _logAction($"Found {elementsInCategory.Count} Generic Model elements in document '{doc.Title}'");
                    }
                }
                catch (Exception ex)
                {
                    _logAction($"Warning: Could not collect category {category} from document: {ex.Message}");
                }
            }

            _logAction($"Collected {mepElements.Count} total elements from document '{doc.Title}'.");
            return mepElements;
        }

        public List<Element> GetRelevantElementsFromCollection(List<Element> elements)
        {
            var relevantElements = new List<Element>();
            int skippedNoCategory = 0;
            int skippedExcluded = 0;
            int skippedViewSpecific = 0;
            int skippedElementType = 0;

            // Use same comprehensive list as GetAllMepElementsFromDocument
            var relevantCategories = new List<BuiltInCategory>
            {
                // Piping
                BuiltInCategory.OST_PipeFitting,
                BuiltInCategory.OST_PipeAccessory,
                BuiltInCategory.OST_PipeCurves,
                BuiltInCategory.OST_FlexPipeCurves,
                BuiltInCategory.OST_PipeInsulations,
                
                // Ductwork
                BuiltInCategory.OST_DuctFitting,
                BuiltInCategory.OST_DuctAccessory,
                BuiltInCategory.OST_DuctCurves,
                BuiltInCategory.OST_FlexDuctCurves,
                BuiltInCategory.OST_DuctTerminal,
                BuiltInCategory.OST_DuctInsulations,
                
                // Mechanical
                BuiltInCategory.OST_MechanicalEquipment,
                BuiltInCategory.OST_PlumbingFixtures,
                BuiltInCategory.OST_Sprinklers,
                BuiltInCategory.OST_FireAlarmDevices,
                
                // Electrical
                BuiltInCategory.OST_ElectricalEquipment,
                BuiltInCategory.OST_ElectricalFixtures,
                BuiltInCategory.OST_LightingFixtures,
                BuiltInCategory.OST_LightingDevices,
                BuiltInCategory.OST_CableTray,
                BuiltInCategory.OST_CableTrayFitting,
                BuiltInCategory.OST_Conduit,
                BuiltInCategory.OST_ConduitFitting,
                BuiltInCategory.OST_DataDevices,
                BuiltInCategory.OST_CommunicationDevices,
                BuiltInCategory.OST_SecurityDevices,
                BuiltInCategory.OST_NurseCallDevices,
                BuiltInCategory.OST_TelephoneDevices,
                
                // Structural
                BuiltInCategory.OST_StructuralFraming,
                BuiltInCategory.OST_StructuralColumns,
                BuiltInCategory.OST_StructuralFoundation,
                BuiltInCategory.OST_StructuralFramingSystem,
                BuiltInCategory.OST_StructuralStiffener,
                BuiltInCategory.OST_StructuralTruss,
                
                // Architectural
                BuiltInCategory.OST_Walls,
                BuiltInCategory.OST_Doors,
                BuiltInCategory.OST_Windows,
                
                // General
                BuiltInCategory.OST_GenericModel,
                BuiltInCategory.OST_SpecialityEquipment,
                BuiltInCategory.OST_Mass,
                BuiltInCategory.OST_Furniture,
                BuiltInCategory.OST_FurnitureSystems,
                BuiltInCategory.OST_Casework,
                BuiltInCategory.OST_Parts
            };

            foreach (var element in elements)
            {
                try
                {
                    // Skip element types
                    if (element is ElementType)
                    {
                        skippedElementType++;
                        continue;
                    }

                    // Skip view-specific elements
                    if (element.OwnerViewId != ElementId.InvalidElementId)
                    {
                        skippedViewSpecific++;
                        continue;
                    }

                    if (element.Category == null)
                    {
                        _logAction($"    Skipping element without category: {element.Id} (non-copyable)");
                        skippedNoCategory++;
                        continue;
                    }

                    var categoryId = (BuiltInCategory)element.Category.Id.IntegerValue;

                    // Check if this is an excluded category
                    if (ExcludedCategories.Contains(categoryId))
                    {
                        skippedExcluded++;
                        continue;
                    }

                    if (relevantCategories.Contains(categoryId))
                    {
                        relevantElements.Add(element);
                    }
                }
                catch (Exception ex)
                {
                    _logAction($"Warning: Error checking element {element.Id}: {ex.Message}");
                }
            }

            // Log summary if significant filtering occurred
            int totalSkipped = skippedNoCategory + skippedExcluded + skippedViewSpecific + skippedElementType;
            if (totalSkipped > 10)
            {
                _logAction($"  Element filtering summary: {relevantElements.Count} relevant, {totalSkipped} skipped");
                if (skippedNoCategory > 0) _logAction($"    - No category: {skippedNoCategory}");
                if (skippedExcluded > 0) _logAction($"    - Excluded categories: {skippedExcluded}");
                if (skippedViewSpecific > 0) _logAction($"    - View-specific: {skippedViewSpecific}");
                if (skippedElementType > 0) _logAction($"    - Element types: {skippedElementType}");
            }

            return relevantElements;
        }

        public HashSet<ElementId> CollectExistingWorksetElements(List<MappingRecord> mapping, Dictionary<string, List<ElementId>> packageGroupMapping, WorksetMapper worksetMapper)
        {
            var existingElements = new HashSet<ElementId>();

            var targetWorksetNames = mapping.Where(m => m.ModelPackageCode != "NO EXPORT")
                                           .Select(m => m.WorksetName)
                                           .ToHashSet(StringComparer.OrdinalIgnoreCase);

            targetWorksetNames.Add("DX_STB");
            targetWorksetNames.Add("DX_ELT");
            targetWorksetNames.Add("DX_RR");
            targetWorksetNames.Add("DX_FND");

            _logAction($"Checking for existing elements in target worksets: {string.Join(", ", targetWorksetNames)}");

            foreach (var worksetName in targetWorksetNames)
            {
                try
                {
                    var workset = new FilteredWorksetCollector(_doc)
                        .OfKind(WorksetKind.UserWorkset)
                        .FirstOrDefault(w => w.Name.Equals(worksetName, StringComparison.OrdinalIgnoreCase));

                    if (workset == null)
                    {
                        _logAction($"  Workset '{worksetName}' not found in model");
                        continue;
                    }

                    var elementsInWorkset = GetAllMepElements()
                        .Where(e => e.WorksetId == workset.Id)
                        .ToList();

                    if (elementsInWorkset.Any())
                    {
                        _logAction($"  Found {elementsInWorkset.Count} existing elements in workset '{worksetName}'");

                        foreach (var element in elementsInWorkset)
                        {
                            existingElements.Add(element.Id);
                        }

                        var mappingRecord = mapping.FirstOrDefault(m =>
                            m.WorksetName.Equals(worksetName, StringComparison.OrdinalIgnoreCase));

                        string packageCode;
                        if (mappingRecord != null)
                        {
                            packageCode = mappingRecord.NormalizedPackageCode;
                        }
                        else
                        {
                            packageCode = worksetMapper.GetIflsCodeForWorkset(worksetName);
                        }

                        if (!packageGroupMapping.ContainsKey(packageCode))
                            packageGroupMapping[packageCode] = new List<ElementId>();

                        packageGroupMapping[packageCode].AddRange(elementsInWorkset.Select(e => e.Id));

                        _logAction($"  Preserved {elementsInWorkset.Count} elements from '{worksetName}' for export as '{packageCode}'");
                    }
                    else
                    {
                        _logAction($"  No elements found in existing workset '{worksetName}'");
                    }
                }
                catch (Exception ex)
                {
                    _logAction($"  Error checking workset '{worksetName}': {ex.Message}");
                }
            }

            _logAction($"Total existing elements preserved: {existingElements.Count}");
            return existingElements;
        }
    }
}