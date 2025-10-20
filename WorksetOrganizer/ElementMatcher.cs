using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using Autodesk.Revit.DB;

namespace WorksetOrchestrator
{
    public class ElementMatcher
    {
        private readonly Action<string> _logAction;
        private readonly ElementClassifier _classifier;

        public ElementMatcher(Action<string> logAction, ElementClassifier classifier)
        {
            _logAction = logAction;
            _classifier = classifier;
        }

        public List<Element> FindMatchingElements(List<Element> elements, MappingRecord record)
        {
            var matchingElements = new List<Element>();
            string systemPattern = record.SystemNameInModel?.Trim();
            string systemDescription = record.SystemDescription?.Trim();

            _logAction($"Searching for elements matching workset '{record.WorksetName}' with pattern '{systemPattern}' in {elements.Count} elements...");

            if (record.WorksetName?.ToUpper() == "DX_ELT")
            {
                _logAction($"Processing DX_ELT workset - checking for electrical elements");

                foreach (var element in elements)
                {
                    if (_classifier.IsElectricalElement(element))
                    {
                        matchingElements.Add(element);
                    }
                }

                _logAction($"Found {matchingElements.Count} electrical elements for DX_ELT");
                return matchingElements;
            }

            if (record.WorksetName?.ToUpper() == "DX_STB")
            {
                _logAction($"Processing DX_STB workset - checking for structural elements including Generic Models");

                foreach (var element in elements)
                {
                    if (_classifier.IsStructuralElement(element) || _classifier.IsGenericModelStructural(element))
                    {
                        matchingElements.Add(element);
                        _logAction($"  Added structural element {element.Id} ({element.Category?.Name ?? "No Category"}) to DX_STB");
                    }
                }

                _logAction($"Found {matchingElements.Count} structural elements (including Generic Models) for DX_STB");
                return matchingElements;
            }

            if (IsWorksetCategoryBased(record.WorksetName))
            {
                _logAction($"Processing category-based workset '{record.WorksetName}' - checking by category and parameters");

                foreach (var element in elements)
                {
                    if (IsElementMatchingByCategory(element, record))
                    {
                        matchingElements.Add(element);
                    }
                }

                _logAction($"Found {matchingElements.Count} category-based elements for {record.WorksetName}");
                return matchingElements;
            }

            if (string.IsNullOrEmpty(systemPattern) || systemPattern == "-")
            {
                _logAction($"No system pattern provided for '{record.WorksetName}' - using category and parameter-based matching");

                foreach (var element in elements)
                {
                    if (IsElementMatchingByCategory(element, record))
                    {
                        matchingElements.Add(element);
                    }
                }

                return matchingElements;
            }

            foreach (var element in elements)
            {
                if (IsElementMatchingSystem(element, systemPattern, systemDescription))
                {
                    matchingElements.Add(element);
                }
            }

            return matchingElements;
        }

        private bool IsWorksetCategoryBased(string worksetName)
        {
            if (string.IsNullOrEmpty(worksetName))
                return false;

            var categoryBasedWorksets = new[]
            {
                "DX_STB",
                "DX_RR",
                "DX_FND",
                "DX_ELT"
            };

            return categoryBasedWorksets.Any(w =>
                w.Equals(worksetName, StringComparison.OrdinalIgnoreCase));
        }

        private bool IsElementMatchingByCategory(Element element, MappingRecord record)
        {
            try
            {
                string worksetName = record.WorksetName?.ToUpper();
                string description = record.SystemDescription?.ToUpper();

                if (worksetName == "DX_ELT")
                {
                    if (_classifier.IsElectricalElement(element))
                    {
                        _logAction($"  MATCH (Category-ELT): Element {element.Id} identified as electrical");
                        return true;
                    }
                }
                else if (worksetName == "DX_STB")
                {
                    if (_classifier.IsStructuralElement(element))
                    {
                        _logAction($"  MATCH (Category-STB): Element {element.Id} identified as structural");
                        return true;
                    }

                    if (element.Category?.Id.IntegerValue == (int)BuiltInCategory.OST_GenericModel)
                    {
                        if (_classifier.IsGenericModelStructural(element))
                        {
                            _logAction($"  MATCH (GenericModel-STB): Generic Model {element.Id} identified as structural");
                            return true;
                        }
                    }
                }
                else if (worksetName == "DX_RR")
                {
                    if (_classifier.IsCleanroomPartition(element))
                    {
                        _logAction($"  MATCH (Category-RR): Element {element.Id} identified as cleanroom partition");
                        return true;
                    }
                }
                else if (worksetName == "DX_FND")
                {
                    if (_classifier.IsFoundation(element))
                    {
                        _logAction($"  MATCH (Category-FND): Element {element.Id} identified as foundation/pedestal");
                        return true;
                    }
                }
                else
                {
                    if (HasMatchingWorksetParameter(element, worksetName, description))
                    {
                        _logAction($"  MATCH (Parameter): Element {element.Id} matched by workset parameter");
                        return true;
                    }
                }

                return false;
            }
            catch (Exception ex)
            {
                _logAction($"Warning: Error in category matching for element {element.Id}: {ex.Message}");
                return false;
            }
        }

        private bool IsElementMatchingSystem(Element element, string systemPattern, string systemDescription)
        {
            try
            {
                if (IsElectricalPattern(systemPattern))
                {
                    if (_classifier.IsElectricalElement(element))
                    {
                        _logAction($"  MATCH (Electrical): Element {element.Id} - Cable Tray/Busbar detected for ELT");
                        return true;
                    }
                }

                var systemValues = new List<string>();

                var parameterIds = new List<BuiltInParameter>
                {
                    BuiltInParameter.RBS_SYSTEM_NAME_PARAM,
                    BuiltInParameter.RBS_SYSTEM_CLASSIFICATION_PARAM,
                    BuiltInParameter.RBS_SYSTEM_ABBREVIATION_PARAM,
                    BuiltInParameter.ELEM_TYPE_PARAM,
                };

                foreach (var paramId in parameterIds)
                {
                    var param = element.get_Parameter(paramId);
                    if (param != null && !string.IsNullOrEmpty(param.AsString()))
                    {
                        systemValues.Add(param.AsString());
                    }
                }

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

                if (element is FamilyInstance familyInstance)
                {
                    try
                    {
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

                systemValues = systemValues.Where(v => !string.IsNullOrEmpty(v)).Distinct().ToList();

                if (systemValues.Count == 0)
                    return false;

                foreach (var systemValue in systemValues)
                {
                    if (DoesSystemValueMatchPattern(systemValue, systemPattern))
                    {
                        _logAction($"  MATCH: Element {element.Id} matched pattern.");
                        return true;
                    }
                }

                if (!string.IsNullOrEmpty(systemDescription))
                {
                    foreach (var systemValue in systemValues)
                    {
                        if (systemValue.IndexOf(systemDescription, StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            _logAction($"  MATCH (Description): Element {element.Id} - contains description.");
                            return true;
                        }
                    }
                }

                return false;
            }
            catch (Exception ex)
            {
                _logAction($"Warning: Error checking element {element.Id}: {ex.Message}");
                return false;
            }
        }

        private bool IsElectricalPattern(string pattern)
        {
            if (string.IsNullOrEmpty(pattern))
                return false;

            var electricalKeywords = new[] { "ELT", "ELECTRICAL", "CABLE", "BUSBAR", "POWER", "LIGHTING" };

            return electricalKeywords.Any(keyword =>
                pattern.IndexOf(keyword, StringComparison.OrdinalIgnoreCase) >= 0);
        }

        private bool HasMatchingWorksetParameter(Element element, string targetWorksetName, string description)
        {
            try
            {
                var parameterSources = new List<Element> { element };

                ElementType elementType = element.Document.GetElement(element.GetTypeId()) as ElementType;
                if (elementType != null)
                    parameterSources.Add(elementType);

                foreach (var paramSource in parameterSources)
                {
                    var worksetParam = paramSource.LookupParameter("Workset");
                    if (worksetParam != null && !string.IsNullOrEmpty(worksetParam.AsString()))
                    {
                        string worksetValue = worksetParam.AsString();

                        string cleanTargetName = targetWorksetName.Replace("DX_", "");
                        if (worksetValue.IndexOf(cleanTargetName, StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            _logAction($"  Element {element.Id} matched by workset parameter: '{worksetValue}' contains '{cleanTargetName}'");
                            return true;
                        }

                        if (!string.IsNullOrEmpty(description))
                        {
                            var descriptionWords = description.Split(new char[] { ' ', ',', ';' }, StringSplitOptions.RemoveEmptyEntries);
                            foreach (var word in descriptionWords)
                            {
                                if (word.Length > 3 && worksetValue.IndexOf(word, StringComparison.OrdinalIgnoreCase) >= 0)
                                {
                                    _logAction($"  Element {element.Id} matched by description word: '{word}' in workset parameter: '{worksetValue}'");
                                    return true;
                                }
                            }
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

        private bool DoesSystemValueMatchPattern(string systemValue, string pattern)
        {
            if (string.IsNullOrEmpty(systemValue) || string.IsNullOrEmpty(pattern))
                return false;

            if (systemValue.Equals(pattern, StringComparison.OrdinalIgnoreCase))
                return true;

            try
            {
                string regexPattern = Regex.Escape(pattern)
                    .Replace("xxx", @"\d{2,3}")
                    .Replace("xx", @"\d{1,3}")
                    .Replace("x", @"\d")
                    .Replace(@"\*", ".*");

                var regex = new Regex(regexPattern, RegexOptions.IgnoreCase);

                if (regex.IsMatch(systemValue))
                    return true;

                string simplifiedPattern = pattern.Replace("xxx", "").Replace("xx", "").Replace("x", "").Trim();
                if (!string.IsNullOrEmpty(simplifiedPattern) &&
                    systemValue.IndexOf(simplifiedPattern, StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    return true;
                }
            }
            catch (Exception ex)
            {
                _logAction($"Regex error for pattern '{pattern}': {ex.Message}");
            }

            string basePattern = pattern.Replace("xxx", "").Replace("xx", "").Replace("x", "").Trim();
            return !string.IsNullOrEmpty(basePattern) &&
                   systemValue.IndexOf(basePattern, StringComparison.OrdinalIgnoreCase) >= 0;
        }
    }
}