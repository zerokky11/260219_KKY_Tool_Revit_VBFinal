using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KKY_Tool_Revit.Models;

namespace KKY_Tool_Revit.Services
{
    public static class ModelParameterExtractionService
    {
        public sealed class ElementParameterValueInfo
        {
            public bool HasParameter { get; set; }
            public string ValueText { get; set; } = string.Empty;
            public string DataTypeToken { get; set; } = string.Empty;
            public double? InternalDoubleValue { get; set; }
        }

        public static int CountExtractableElements(Document doc)
        {
            if (doc == null) throw new ArgumentNullException(nameof(doc));
            return CollectExtractableElements(doc).Count;
        }

        public static IDictionary<string, int> GetExtractableElementSignatureCounts(Document doc)
        {
            if (doc == null) throw new ArgumentNullException(nameof(doc));

            var result = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            foreach (Element element in CollectExtractableElements(doc))
            {
                string signature = BuildExtractableElementSignature(doc, element);
                if (string.IsNullOrWhiteSpace(signature))
                {
                    signature = "(Unknown)";
                }

                int currentCount;
                result.TryGetValue(signature, out currentCount);
                result[signature] = currentCount + 1;
            }

            return result;
        }

        public static string BuildReductionSummary(IDictionary<string, int> beforeCounts, IDictionary<string, int> afterCounts, int maxItems = 8)
        {
            if (beforeCounts == null || beforeCounts.Count == 0)
            {
                return string.Empty;
            }

            var removedItems = new List<KeyValuePair<string, int>>();
            foreach (KeyValuePair<string, int> item in beforeCounts)
            {
                int afterCount = 0;
                if (afterCounts != null)
                {
                    afterCounts.TryGetValue(item.Key, out afterCount);
                }

                int removedCount = item.Value - afterCount;
                if (removedCount > 0)
                {
                    removedItems.Add(new KeyValuePair<string, int>(item.Key, removedCount));
                }
            }

            if (removedItems.Count == 0)
            {
                return string.Empty;
            }

            List<KeyValuePair<string, int>> topItems = removedItems
                .OrderByDescending(x => x.Value)
                .ThenBy(x => x.Key, StringComparer.OrdinalIgnoreCase)
                .Take(Math.Max(1, maxItems))
                .ToList();

            string summary = string.Join("; ", topItems.Select(x => x.Key + " " + x.Value + "개"));
            int remainingItemCount = removedItems.Count - topItems.Count;
            if (remainingItemCount > 0)
            {
                summary += " 외 " + remainingItemCount + "개 항목";
            }

            return summary;
        }

        public static IList<Element> GetExtractableElements(Document doc)
        {
            if (doc == null) throw new ArgumentNullException(nameof(doc));
            return CollectExtractableElements(doc);
        }

        public static string GetElementCategoryName(Element element)
        {
            return GetCategoryName(element);
        }

        public static string GetElementFamilyName(Document doc, Element element)
        {
            return GetFamilyName(doc, element);
        }

        public static string GetElementTypeName(Document doc, Element element)
        {
            return GetTypeName(doc, element);
        }

        public static string GetElementParameterValue(Document doc, Element element, string parameterName)
        {
            return GetParameterValue(doc, element, parameterName);
        }

        public static ElementParameterValueInfo GetElementParameterValueInfo(Document doc, Element element, string parameterName)
        {
            return GetParameterValueInfo(doc, element, parameterName);
        }

        public static bool HasElementParameter(Document doc, Element element, string parameterName)
        {
            return FindParameterOnElementOrType(doc, element, parameterName) != null;
        }

        public static string ExportModelParameters(UIApplication uiapp, BatchPrepareSession session, string outputFolder, string parameterNamesCsv, Action<string> log)
        {
            if (session == null) throw new ArgumentNullException(nameof(session));
            return ExportModelParameters(uiapp, session.CleanedOutputPaths, string.IsNullOrWhiteSpace(outputFolder) ? session.OutputFolder : outputFolder, parameterNamesCsv, log);
        }

        public static string ExportModelParameters(UIApplication uiapp, IEnumerable<string> targetPaths, string outputFolder, string parameterNamesCsv, Action<string> log)
        {
            if (uiapp == null) throw new ArgumentNullException(nameof(uiapp));

            List<string> paths = (targetPaths ?? Enumerable.Empty<string>())
                .Where(x => !string.IsNullOrWhiteSpace(x) && File.Exists(x))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();
            if (paths.Count == 0) throw new InvalidOperationException("추출 대상 모델 파일이 없습니다.");

            List<string> parameterNames = SplitParameterNames(parameterNamesCsv);
            if (parameterNames.Count == 0) throw new InvalidOperationException("추출할 파라미터명이 없습니다. 하나 이상 입력해 주세요.");

            if (string.IsNullOrWhiteSpace(outputFolder))
            {
                outputFolder = Path.GetDirectoryName(paths[0]);
            }
            Directory.CreateDirectory(outputFolder);

            string csvPath = Path.Combine(outputFolder, "ModelParameterExport_" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + ".csv");
            var sb = new StringBuilder();
            var header = new List<string> { "FileName", "ElementId", "Category", "FamilyName", "TypeName" };
            header.AddRange(parameterNames);
            sb.AppendLine(string.Join(",", header.Select(Csv)));

            foreach (string path in paths)
            {
                Document doc = null;
                try
                {
                    log?.Invoke("모델 파라미터 추출 시작: " + path);
                    ModelPath modelPath = ModelPathUtils.ConvertUserVisiblePathToModelPath(path);
                    doc = uiapp.Application.OpenDocumentFile(modelPath, new OpenOptions());
                    IList<Element> elements = CollectExtractableElements(doc);

                    int rowCount = 0;
                    string fileName = Path.GetFileName(path);
                    foreach (Element element in elements)
                    {
                        var row = new List<string>();
                        row.Add(fileName);
                        row.Add(element.Id.IntegerValue.ToString());
                        row.Add(GetCategoryName(element));
                        row.Add(GetFamilyName(doc, element));
                        row.Add(GetTypeName(doc, element));
                        foreach (string parameterName in parameterNames)
                        {
                            row.Add(GetParameterValue(doc, element, parameterName));
                        }
                        sb.AppendLine(string.Join(",", row.Select(Csv)));
                        rowCount++;
                    }

                    log?.Invoke("모델 파라미터 추출 완료: " + fileName + " / 행 " + rowCount);
                }
                finally
                {
                    try
                    {
                        if (doc != null && doc.IsValidObject)
                        {
                            doc.Close(false);
                        }
                    }
                    catch
                    {
                    }
                }
            }

            File.WriteAllText(csvPath, sb.ToString(), new UTF8Encoding(true));
            log?.Invoke("모델 파라미터 CSV 저장: " + csvPath);
            return csvPath;
        }

        private static IList<Element> CollectExtractableElements(Document doc)
        {
            HashSet<int> schedulableCategoryIds = GetSchedulableCategoryIds(doc);
            return new FilteredElementCollector(doc)
                .WhereElementIsNotElementType()
                .Cast<Element>()
                .Where(x => x != null)
                .Where(x => IsEligibleModelElement(x, schedulableCategoryIds))
                .OrderBy(x => x.Id.IntegerValue)
                .ToList();
        }

        private static string BuildExtractableElementSignature(Document doc, Element element)
        {
            string categoryName = GetCategoryName(element);
            string familyName = GetFamilyName(doc, element);
            string typeName = GetTypeName(doc, element);

            var parts = new List<string>();
            if (!string.IsNullOrWhiteSpace(categoryName))
            {
                parts.Add(categoryName.Trim());
            }

            if (!string.IsNullOrWhiteSpace(familyName) &&
                !parts.Any(x => string.Equals(x, familyName.Trim(), StringComparison.OrdinalIgnoreCase)))
            {
                parts.Add(familyName.Trim());
            }

            if (!string.IsNullOrWhiteSpace(typeName) &&
                !parts.Any(x => string.Equals(x, typeName.Trim(), StringComparison.OrdinalIgnoreCase)))
            {
                parts.Add(typeName.Trim());
            }

            return string.Join(" | ", parts.Where(x => !string.IsNullOrWhiteSpace(x)));
        }

        private static bool IsEligibleModelElement(Element element, ISet<int> schedulableCategoryIds)
        {
            if (element == null) return false;
            if (element.ViewSpecific) return false;
            if (element.Category == null) return false;
            if (element.Category.CategoryType != CategoryType.Model) return false;
            if (element is View) return false;
            if (element is ReferencePlane) return false;
            if (element is CurveElement) return false;
            if (element is Grid) return false;
            if (element is Level) return false;
            if (element is Group) return false;
            if (element is AssemblyInstance) return false;
            if (element is Room) return false;
            if (element is Area) return false;
            if (element is MEPSystem) return false;

            int categoryId;
            try
            {
                categoryId = element.Category.Id != null ? element.Category.Id.IntegerValue : 0;
            }
            catch
            {
                return false;
            }

            if (schedulableCategoryIds == null || !schedulableCategoryIds.Contains(categoryId)) return false;
            return !IsExplicitlyExcludedCategory(element.Category);
        }

        private static HashSet<int> GetSchedulableCategoryIds(Document doc)
        {
            var result = new HashSet<int>();
            if (doc == null) return result;

            Categories categories = null;
            try { categories = doc.Settings != null ? doc.Settings.Categories : null; } catch { }
            if (categories == null) return result;

            foreach (Category category in categories)
            {
                if (category == null || category.Id == null) continue;

                int categoryId;
                try { categoryId = category.Id.IntegerValue; }
                catch { continue; }

                if (categoryId == 0) continue;
                if (IsSchedulableCategory(doc, category))
                {
                    result.Add(categoryId);
                }
            }

            return result;
        }

        private static bool IsSchedulableCategory(Document doc, Category category)
        {
            if (doc == null || category == null || category.Id == null) return false;
            if (category.CategoryType != CategoryType.Model) return false;
            if (IsExplicitlyIncludedCategory(category)) return true;
            if (IsExplicitlyExcludedCategory(category)) return false;

            bool? directCheck = TryIsValidCategoryForSchedule(doc, category);
            if (directCheck.HasValue) return directCheck.Value;

            HashSet<int> validIds = TryGetValidCategoriesForSchedule(doc);
            if (validIds != null && validIds.Count > 0)
            {
                return validIds.Contains(category.Id.IntegerValue);
            }

            return true;
        }

        private static bool IsExplicitlyIncludedCategory(Category category)
        {
            return MatchesBuiltInCategoryNames(category,
                "OST_Curtain_Systems",
                "OST_CurtaSystem");
        }

        private static bool IsExplicitlyExcludedCategory(Category category)
        {
            if (category == null) return true;
            if (IsExplicitlyIncludedCategory(category)) return false;

            string name = string.Empty;
            try { name = category.Name ?? string.Empty; } catch { }
            string normalized = name.Trim();
            if (string.IsNullOrWhiteSpace(normalized)) return true;

            string[] blockedKeywords =
            {
                "Analytical",
                "Load",
                "Placeholder",
                "Zone",
                "Area",
                "Grid",
                "Level",
                "Reference",
                "Center Line",
                "Centerline",
                "Annotation",
                "Space",
                "System",
                "Material",
                "Project Information",
                "Sun Path",
                "Pipe Segment",
                "Primary Contour",
                "Legend Component",
                "Systems",
                "Boundary",
                "Separation"
            };

            if (blockedKeywords.Any(keyword => normalized.IndexOf(keyword, StringComparison.OrdinalIgnoreCase) >= 0))
            {
                return true;
            }

            return MatchesBuiltInCategoryNames(category,
                "OST_AnalyticalNodes",
                "OST_AnalyticalLinks",
                "OST_AnalyticalPipeNodes",
                "OST_AnalyticalPipeConnections",
                "OST_AnalyticalSpaces",
                "OST_GridChains",
                "OST_Grids",
                "OST_Levels",
                "OST_Rooms",
                "OST_Areas",
                "OST_Lines",
                "OST_CLines",
                "OST_IOSModelGroups",
                "OST_Assemblies",
                "OST_MEPSpaces",
                "OST_HVAC_Zones",
                "OST_AreaSchemeLines",
                "OST_RoomSeparationLines",
                "OST_MEPAnalyticalAirLoop",
                "OST_MEPAnalyticalWaterLoop",
                "OST_ElectricalLoadAreas",
                "OST_ElectricalLoadClassifications",
                "OST_LoadCases",
                "OST_LoadCombinations",
                "OST_Loads",
                "OST_PointLoadTags",
                "OST_LineLoadTags",
                "OST_AreaLoadTags",
                "OST_PlaceHolderDucts",
                "OST_PlaceHolderPipes",
                "OST_PlaceHolderCableTray",
                "OST_PlaceHolderConduits",
                "OST_ProjectInformation",
                "OST_SunPath",
                "OST_PipeSegments",
                "OST_PrimaryContour",
                "OST_LegendComponents",
                "OST_Materials");
        }

        private static bool MatchesBuiltInCategoryNames(Category category, params string[] builtInCategoryNames)
        {
            if (category == null || category.Id == null || builtInCategoryNames == null || builtInCategoryNames.Length == 0)
            {
                return false;
            }

            int categoryId;
            try
            {
                categoryId = category.Id.IntegerValue;
            }
            catch
            {
                return false;
            }

            foreach (string name in builtInCategoryNames)
            {
                if (string.IsNullOrWhiteSpace(name)) continue;

                try
                {
                    if (Enum.TryParse(name, out BuiltInCategory builtInCategory) && categoryId == (int)builtInCategory)
                    {
                        return true;
                    }
                }
                catch
                {
                }
            }

            return false;
        }

        private static bool? TryIsValidCategoryForSchedule(Document doc, Category category)
        {
            foreach (MethodInfo method in typeof(ViewSchedule).GetMethods(BindingFlags.Public | BindingFlags.Static)
                         .Where(x => string.Equals(x.Name, "IsValidCategoryForSchedule", StringComparison.Ordinal)))
            {
                ParameterInfo[] parameters = method.GetParameters();
                try
                {
                    if (parameters.Length == 1)
                    {
                        if (parameters[0].ParameterType == typeof(ElementId))
                        {
                            return Convert.ToBoolean(method.Invoke(null, new object[] { category.Id }));
                        }

                        if (parameters[0].ParameterType == typeof(Category))
                        {
                            return Convert.ToBoolean(method.Invoke(null, new object[] { category }));
                        }
                    }
                    else if (parameters.Length == 2)
                    {
                        if (parameters[0].ParameterType == typeof(Document) && parameters[1].ParameterType == typeof(ElementId))
                        {
                            return Convert.ToBoolean(method.Invoke(null, new object[] { doc, category.Id }));
                        }

                        if (parameters[0].ParameterType == typeof(Document) && parameters[1].ParameterType == typeof(Category))
                        {
                            return Convert.ToBoolean(method.Invoke(null, new object[] { doc, category }));
                        }
                    }
                }
                catch
                {
                }
            }

            return null;
        }

        private static HashSet<int> TryGetValidCategoriesForSchedule(Document doc)
        {
            foreach (MethodInfo method in typeof(ViewSchedule).GetMethods(BindingFlags.Public | BindingFlags.Static)
                         .Where(x => string.Equals(x.Name, "GetValidCategoriesForSchedule", StringComparison.Ordinal)))
            {
                ParameterInfo[] parameters = method.GetParameters();
                try
                {
                    object raw = null;
                    if (parameters.Length == 0)
                    {
                        raw = method.Invoke(null, null);
                    }
                    else if (parameters.Length == 1 && parameters[0].ParameterType == typeof(Document))
                    {
                        raw = method.Invoke(null, new object[] { doc });
                    }

                    HashSet<int> categoryIds = ExtractCategoryIds(raw);
                    if (categoryIds.Count > 0) return categoryIds;
                }
                catch
                {
                }
            }

            return null;
        }

        private static HashSet<int> ExtractCategoryIds(object raw)
        {
            var result = new HashSet<int>();
            var enumerable = raw as System.Collections.IEnumerable;
            if (enumerable == null || raw is string) return result;

            foreach (object item in enumerable)
            {
                if (item == null) continue;

                if (item is ElementId elementId)
                {
                    result.Add(elementId.IntegerValue);
                    continue;
                }

                if (item is Category category && category.Id != null)
                {
                    result.Add(category.Id.IntegerValue);
                    continue;
                }

                PropertyInfo idProperty = item.GetType().GetProperty("Id", BindingFlags.Public | BindingFlags.Instance);
                if (idProperty == null) continue;

                object idValue = null;
                try { idValue = idProperty.GetValue(item, null); } catch { }

                if (idValue is ElementId reflectedId)
                {
                    result.Add(reflectedId.IntegerValue);
                }
                else if (idValue is Category reflectedCategory && reflectedCategory.Id != null)
                {
                    result.Add(reflectedCategory.Id.IntegerValue);
                }
            }

            return result;
        }

        private static List<string> SplitParameterNames(string csv)
        {
            if (string.IsNullOrWhiteSpace(csv)) return new List<string>();
            return csv.Split(new[] { ',', ';', '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries)
                .Select(x => x.Trim())
                .Where(x => !string.IsNullOrWhiteSpace(x))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        private static string GetCategoryName(Element element)
        {
            try { return element.Category != null ? element.Category.Name : string.Empty; } catch { return string.Empty; }
        }

        private static string GetFamilyName(Document doc, Element element)
        {
            if (doc == null || element == null) return string.Empty;
            try
            {
                FamilyInstance familyInstance = element as FamilyInstance;
                if (familyInstance != null && familyInstance.Symbol != null && familyInstance.Symbol.Family != null)
                {
                    return familyInstance.Symbol.Family.Name ?? string.Empty;
                }
            }
            catch
            {
            }

            try
            {
                ElementType type = GetElementType(doc, element);
                if (type != null)
                {
                    string familyName = type.FamilyName;
                    if (!string.IsNullOrWhiteSpace(familyName)) return familyName;
                }
            }
            catch
            {
            }

            return string.Empty;
        }

        private static string GetTypeName(Document doc, Element element)
        {
            if (doc == null || element == null) return string.Empty;
            try
            {
                ElementType type = GetElementType(doc, element);
                if (type != null)
                {
                    return type.Name ?? string.Empty;
                }
            }
            catch
            {
            }
            return string.Empty;
        }

        private static ElementType GetElementType(Document doc, Element element)
        {
            if (doc == null || element == null) return null;
            try
            {
                ElementId typeId = element.GetTypeId();
                if (typeId == null || typeId == ElementId.InvalidElementId) return null;
                return doc.GetElement(typeId) as ElementType;
            }
            catch
            {
                return null;
            }
        }

        private static string GetParameterValue(Document doc, Element element, string parameterName)
        {
            return GetParameterValueInfo(doc, element, parameterName).ValueText ?? string.Empty;
        }

        private static ElementParameterValueInfo GetParameterValueInfo(Document doc, Element element, string parameterName)
        {
            var info = new ElementParameterValueInfo();
            Parameter parameter = FindParameterOnElementOrType(doc, element, parameterName);
            if (parameter == null) return info;

            info.HasParameter = true;
            info.DataTypeToken = GetDefinitionDataTypeToken(parameter.Definition);
            if (parameter.StorageType == StorageType.Double)
            {
                try
                {
                    info.InternalDoubleValue = parameter.AsDouble();
                }
                catch
                {
                    info.InternalDoubleValue = null;
                }
            }

            try
            {
                string valueString = parameter.AsValueString();
                if (!string.IsNullOrWhiteSpace(valueString))
                {
                    info.ValueText = valueString;
                    return info;
                }
            }
            catch
            {
            }

            try
            {
                switch (parameter.StorageType)
                {
                    case StorageType.String:
                        info.ValueText = parameter.AsString() ?? string.Empty;
                        break;
                    case StorageType.Integer:
                        info.ValueText = parameter.AsInteger().ToString();
                        break;
                    case StorageType.Double:
                        info.ValueText = parameter.AsDouble().ToString(System.Globalization.CultureInfo.InvariantCulture);
                        break;
                    case StorageType.ElementId:
                        ElementId id = parameter.AsElementId();
                        if (id == null || id == ElementId.InvalidElementId)
                        {
                            info.ValueText = string.Empty;
                            break;
                        }
                        Element refElement = doc.GetElement(id);
                        info.ValueText = refElement != null ? (refElement.Name ?? id.IntegerValue.ToString()) : id.IntegerValue.ToString();
                        break;
                    default:
                        info.ValueText = string.Empty;
                        break;
                }
            }
            catch
            {
                info.ValueText = string.Empty;
            }

            return info;
        }

        private static Parameter FindParameterOnElementOrType(Document doc, Element element, string parameterName)
        {
            if (doc == null || element == null || string.IsNullOrWhiteSpace(parameterName)) return null;
            Parameter parameter = FindParameterByName(element, parameterName);
            if (parameter != null) return parameter;

            ElementType type = GetElementType(doc, element);
            if (type != null)
            {
                parameter = FindParameterByName(type, parameterName);
                if (parameter != null) return parameter;
            }

            foreach (BuiltInParameter builtInParameter in GetBuiltInParameterCandidates(parameterName))
            {
                parameter = TryGetBuiltInParameter(element, builtInParameter);
                if (parameter != null) return parameter;

                if (type != null)
                {
                    parameter = TryGetBuiltInParameter(type, builtInParameter);
                    if (parameter != null) return parameter;
                }
            }

            return null;
        }

        private static Parameter FindParameterByName(Element owner, string parameterName)
        {
            if (owner == null || string.IsNullOrWhiteSpace(parameterName)) return null;
            try
            {
                Parameter direct = owner.LookupParameter(parameterName);
                if (IsUsableNamedParameter(direct, parameterName)) return direct;
            }
            catch
            {
            }

            try
            {
                IList<Parameter> directMatches = owner.GetParameters(parameterName);
                if (directMatches != null)
                {
                    foreach (Parameter parameter in directMatches)
                    {
                        if (IsUsableNamedParameter(parameter, parameterName)) return parameter;
                    }
                }
            }
            catch
            {
            }

            try
            {
                foreach (Parameter parameter in owner.Parameters)
                {
                    if (IsUsableNamedParameter(parameter, parameterName)) return parameter;
                }
            }
            catch
            {
            }

            return null;
        }

        private static bool IsUsableNamedParameter(Parameter parameter, string expectedName)
        {
            if (parameter == null || parameter.Definition == null) return false;
            string actualName = parameter.Definition.Name;
            if (string.IsNullOrWhiteSpace(actualName)) return false;

            string normalizedActual = NormalizeParameterToken(actualName);
            string normalizedExpected = NormalizeParameterToken(expectedName);
            return !string.IsNullOrWhiteSpace(normalizedActual)
                   && string.Equals(normalizedActual, normalizedExpected, StringComparison.OrdinalIgnoreCase);
        }

        private static Parameter TryGetBuiltInParameter(Element owner, BuiltInParameter builtInParameter)
        {
            if (owner == null) return null;

            try
            {
                Parameter parameter = owner.get_Parameter(builtInParameter);
                return parameter != null && parameter.Definition != null ? parameter : null;
            }
            catch
            {
                return null;
            }
        }

        private static IEnumerable<BuiltInParameter> GetBuiltInParameterCandidates(string parameterName)
        {
            string token = NormalizeParameterToken(parameterName);
            if (string.IsNullOrWhiteSpace(token)) yield break;

            if (IsAnyToken(token, "area", "면적"))
            {
                yield return BuiltInParameter.HOST_AREA_COMPUTED;
                yield return BuiltInParameter.ROOM_AREA;
            }

            if (IsAnyToken(token, "volume", "vol", "부피", "볼륨"))
            {
                yield return BuiltInParameter.HOST_VOLUME_COMPUTED;
                yield return BuiltInParameter.ROOM_VOLUME;
            }

            if (IsAnyToken(token, "length", "len", "길이"))
            {
                yield return BuiltInParameter.CURVE_ELEM_LENGTH;
            }
        }

        private static bool IsAnyToken(string token, params string[] candidates)
        {
            if (string.IsNullOrWhiteSpace(token) || candidates == null) return false;
            return candidates.Any(candidate => string.Equals(token, NormalizeParameterToken(candidate), StringComparison.OrdinalIgnoreCase));
        }

        private static string NormalizeParameterToken(string value)
        {
            if (string.IsNullOrWhiteSpace(value)) return string.Empty;

            var buffer = new StringBuilder(value.Length);
            foreach (char ch in value.Trim())
            {
                if (char.IsWhiteSpace(ch) || ch == '_' || ch == '-' || ch == '(' || ch == ')' || ch == '[' || ch == ']')
                {
                    continue;
                }

                buffer.Append(char.ToLowerInvariant(ch));
            }

            return buffer.ToString();
        }

        private static string GetDefinitionDataTypeToken(Definition definition)
        {
            if (definition == null) return string.Empty;

            try
            {
                MethodInfo method = definition.GetType().GetMethod("GetDataType", Type.EmptyTypes);
                if (method != null)
                {
                    object dataType = method.Invoke(definition, null);
                    if (dataType != null)
                    {
                        PropertyInfo typeIdProperty = dataType.GetType().GetProperty("TypeId");
                        if (typeIdProperty != null)
                        {
                            string typeId = typeIdProperty.GetValue(dataType, null) as string;
                            if (!string.IsNullOrWhiteSpace(typeId)) return typeId;
                        }

                        string raw = dataType.ToString();
                        if (!string.IsNullOrWhiteSpace(raw)) return raw;
                    }
                }
            }
            catch
            {
            }

            try
            {
                PropertyInfo parameterTypeProperty = definition.GetType().GetProperty("ParameterType");
                if (parameterTypeProperty != null)
                {
                    object parameterType = parameterTypeProperty.GetValue(definition, null);
                    if (parameterType != null) return parameterType.ToString() ?? string.Empty;
                }
            }
            catch
            {
            }

            return string.Empty;
        }
        private static string Csv(string value)
        {
            string text = value ?? string.Empty;
            bool mustQuote = text.Contains(",") || text.Contains("\"") || text.Contains("\r") || text.Contains("\n");
            if (text.Contains("\""))
            {
                text = text.Replace("\"", "\"\"");
            }

            return mustQuote ? "\"" + text + "\"" : text;
        }
    }
}
