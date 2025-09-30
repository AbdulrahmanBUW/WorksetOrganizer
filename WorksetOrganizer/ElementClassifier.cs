using System;
using System.Linq;
using Autodesk.Revit.DB;

namespace WorksetOrchestrator
{
    public class ElementClassifier
    {
        private readonly Action<string> _logAction;

        public ElementClassifier(Action<string> logAction)
        {
            _logAction = logAction;
        }

        public bool IsElectricalElement(Element element)
        {
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
                    if (categoryId == BuiltInCategory.OST_CableTray || categoryId == BuiltInCategory.OST_CableTrayFitting)
                    {
                        return CheckCableTrayWorksetParameter(element);
                    }
                    return true;
                }
            }

            return false;
        }

        public bool IsStructuralElement(Element element)
        {
            _logAction($"  Checking element {element.Id} for structural type - Category: {element.Category?.Name ?? "NULL"}");

            var structuralCategories = new[]
            {
                BuiltInCategory.OST_StructuralFraming,
                BuiltInCategory.OST_StructuralColumns,
                BuiltInCategory.OST_StructuralFoundation,
                BuiltInCategory.OST_StructuralFramingSystem,
                BuiltInCategory.OST_StructuralStiffener,
                BuiltInCategory.OST_StructuralTruss,
                BuiltInCategory.OST_GenericModel
            };

            var elementCategory = element.Category;
            if (elementCategory != null)
            {
                var categoryId = (BuiltInCategory)elementCategory.Id.IntegerValue;
                _logAction($"    Category ID: {categoryId}");

                if (structuralCategories.Contains(categoryId))
                {
                    if (categoryId != BuiltInCategory.OST_GenericModel)
                    {
                        _logAction($"    Pure structural category match: {categoryId}");
                        return CheckStructuralWorksetParameter(element);
                    }
                    else
                    {
                        _logAction($"    Generic Model - checking if structural");
                        return IsGenericModelStructural(element);
                    }
                }
            }

            _logAction($"    No structural category match for element {element.Id}");
            return false;
        }

        public bool IsGenericModelStructural(Element element)
        {
            try
            {
                if (element.Category?.Id.IntegerValue != (int)BuiltInCategory.OST_GenericModel)
                    return false;

                _logAction($"Checking Generic Model {element.Id} for structural classification");

                var parameterSources = new System.Collections.Generic.List<Element> { element };

                ElementType elementType = element.Document.GetElement(element.GetTypeId()) as ElementType;
                if (elementType != null)
                    parameterSources.Add(elementType);

                foreach (var paramSource in parameterSources)
                {
                    var worksetParam = paramSource.LookupParameter("Workset");
                    if (worksetParam != null && !string.IsNullOrEmpty(worksetParam.AsString()))
                    {
                        string worksetValue = worksetParam.AsString();
                        _logAction($"  Generic Model {element.Id} has Workset parameter: '{worksetValue}'");

                        var structuralKeywords = new[]
                        {
                            "Steel", "Structure", "Structural", "STB", "Frame", "Column", "Beam", "Foundation"
                        };

                        bool isStructural = structuralKeywords.Any(keyword =>
                            worksetValue.IndexOf(keyword, StringComparison.OrdinalIgnoreCase) >= 0);

                        if (isStructural)
                        {
                            _logAction($"  Generic Model {element.Id} identified as structural based on Workset parameter: '{worksetValue}'");
                            return true;
                        }
                    }

                    if (element is FamilyInstance familyInstance)
                    {
                        string familyName = familyInstance.Symbol?.FamilyName ?? "";
                        string typeName = familyInstance.Symbol?.Name ?? "";

                        var structuralKeywords = new[]
                        {
                            "Steel", "Structure", "Structural", "STB", "Frame", "Column", "Beam", "Foundation"
                        };

                        bool isFamilyStructural = structuralKeywords.Any(keyword =>
                            familyName.IndexOf(keyword, StringComparison.OrdinalIgnoreCase) >= 0 ||
                            typeName.IndexOf(keyword, StringComparison.OrdinalIgnoreCase) >= 0);

                        if (isFamilyStructural)
                        {
                            _logAction($"  Generic Model {element.Id} identified as structural based on family/type name: '{familyName}' / '{typeName}'");
                            return true;
                        }
                    }
                }

                return false;
            }
            catch (Exception ex)
            {
                _logAction($"Warning: Error checking if Generic Model {element.Id} is structural: {ex.Message}");
                return false;
            }
        }

        public bool IsCleanroomPartition(Element element)
        {
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
                    return CheckParameterForKeywords(element, new[] { "Cleanroom", "Partition", "RR", "Clean Room", "Wall" });
                }
            }

            return false;
        }

        public bool IsFoundation(Element element)
        {
            var foundationCategories = new[]
            {
                BuiltInCategory.OST_StructuralFoundation,
                BuiltInCategory.OST_GenericModel,
                BuiltInCategory.OST_MechanicalEquipment
            };

            var elementCategory = element.Category;
            if (elementCategory != null)
            {
                var categoryId = (BuiltInCategory)elementCategory.Id.IntegerValue;
                if (foundationCategories.Contains(categoryId))
                {
                    return CheckParameterForKeywords(element, new[] { "Foundation", "Pedestal", "FND", "Tool", "Base" });
                }
            }

            return false;
        }

        private bool CheckCableTrayWorksetParameter(Element element)
        {
            try
            {
                ElementType elementType = element.Document.GetElement(element.GetTypeId()) as ElementType;
                if (elementType == null)
                    return true;

                var worksetParam = elementType.LookupParameter("Workset");
                if (worksetParam != null && !string.IsNullOrEmpty(worksetParam.AsString()))
                {
                    string worksetValue = worksetParam.AsString();
                    _logAction($"  Cable Tray/Fitting {element.Id} has Workset parameter: '{worksetValue}'");

                    var nonElectricalKeywords = new[] { "Fire", "HVAC", "Plumbing", "NOT ELECTRICAL" };

                    bool isNonElectrical = nonElectricalKeywords.Any(keyword =>
                        worksetValue.IndexOf(keyword, StringComparison.OrdinalIgnoreCase) >= 0);

                    return !isNonElectrical;
                }

                _logAction($"  Cable Tray/Fitting {element.Id} - no Workset parameter found, defaulting to electrical");
                return true;
            }
            catch (Exception ex)
            {
                _logAction($"Warning: Error checking cable tray workset parameter for element {element.Id}: {ex.Message}");
                return true;
            }
        }

        private bool CheckStructuralWorksetParameter(Element element)
        {
            try
            {
                _logAction($"    Checking structural workset parameter for element {element.Id}");

                ElementType elementType = element.Document.GetElement(element.GetTypeId()) as ElementType;

                if (elementType != null)
                {
                    var worksetParam = elementType.LookupParameter("Workset");
                    if (worksetParam != null && !string.IsNullOrEmpty(worksetParam.AsString()))
                    {
                        string worksetValue = worksetParam.AsString();
                        _logAction($"      Type Workset parameter: '{worksetValue}'");

                        var structuralKeywords = new[]
                        {
                            "Steel", "Structure", "Structural", "STB", "Frame", "Column", "Beam", "Foundation"
                        };

                        bool isStructural = structuralKeywords.Any(keyword =>
                            worksetValue.IndexOf(keyword, StringComparison.OrdinalIgnoreCase) >= 0);

                        if (isStructural)
                        {
                            _logAction($"      MATCH: Found structural keyword in type workset parameter");
                            return true;
                        }
                    }
                    else
                    {
                        _logAction($"      No type Workset parameter found");
                    }
                }

                var instanceWorksetParam = element.LookupParameter("Workset");
                if (instanceWorksetParam != null && !string.IsNullOrEmpty(instanceWorksetParam.AsString()))
                {
                    string worksetValue = instanceWorksetParam.AsString();
                    _logAction($"      Instance Workset parameter: '{worksetValue}'");

                    var structuralKeywords = new[] { "Steel", "Structure", "Structural", "STB", "Frame", "Column", "Beam", "Foundation" };

                    bool isStructural = structuralKeywords.Any(keyword =>
                        worksetValue.IndexOf(keyword, StringComparison.OrdinalIgnoreCase) >= 0);

                    if (isStructural)
                    {
                        _logAction($"      MATCH: Found structural keyword in instance workset parameter");
                        return true;
                    }
                }
                else
                {
                    _logAction($"      No instance Workset parameter found");
                }

                var elementCategory = element.Category;
                if (elementCategory != null)
                {
                    var categoryId = (BuiltInCategory)elementCategory.Id.IntegerValue;
                    var pureStructuralCategories = new[]
                    {
                        BuiltInCategory.OST_StructuralFraming,
                        BuiltInCategory.OST_StructuralColumns,
                        BuiltInCategory.OST_StructuralFoundation,
                        BuiltInCategory.OST_StructuralFramingSystem,
                        BuiltInCategory.OST_StructuralStiffener,
                        BuiltInCategory.OST_StructuralTruss
                    };

                    if (pureStructuralCategories.Contains(categoryId))
                    {
                        _logAction($"      Pure structural category, defaulting to true: {categoryId}");
                        return true;
                    }
                }

                _logAction($"      No structural match found for element {element.Id}");
                return false;
            }
            catch (Exception ex)
            {
                _logAction($"Warning: Error checking structural workset parameter for element {element.Id}: {ex.Message}");
                return false;
            }
        }

        private bool CheckParameterForKeywords(Element element, string[] keywords)
        {
            try
            {
                var parameterSources = new System.Collections.Generic.List<Element> { element };

                ElementType elementType = element.Document.GetElement(element.GetTypeId()) as ElementType;
                if (elementType != null)
                    parameterSources.Add(elementType);

                foreach (var paramSource in parameterSources)
                {
                    var worksetParam = paramSource.LookupParameter("Workset");
                    if (worksetParam != null && !string.IsNullOrEmpty(worksetParam.AsString()))
                    {
                        string worksetValue = worksetParam.AsString();

                        bool matches = keywords.Any(keyword =>
                            worksetValue.IndexOf(keyword, StringComparison.OrdinalIgnoreCase) >= 0);

                        if (matches)
                        {
                            _logAction($"  Element {element.Id} matched workset parameter: '{worksetValue}'");
                            return true;
                        }
                    }
                }

                return false;
            }
            catch (Exception ex)
            {
                _logAction($"Warning: Error checking workset parameter for element {element.Id}: {ex.Message}");
                return false;
            }
        }
    }
}