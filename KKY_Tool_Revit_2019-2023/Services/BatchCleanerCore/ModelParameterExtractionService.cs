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
        public static int CountExtractableElements(Document doc)
        {
            if (doc == null) throw new ArgumentNullException(nameof(doc));
            return CollectExtractableElements(doc).Count;
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
            if (paths.Count == 0) throw new InvalidOperationException("?곕뗄??????뵬????곷뮸??덈뼄.");

            List<string> parameterNames = SplitParameterNames(parameterNamesCsv);
            if (parameterNames.Count == 0) throw new InvalidOperationException("?곕뗄??????뵬沃섎챸苑???已????롪돌 ??곴맒 ??낆젾??곷튊 ??몃빍??");

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
                    log?.Invoke("??욧쉐揶??곕뗄?????뵬 ??용┛: " + path);
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

                    log?.Invoke("??욧쉐揶??곕뗄???袁⑥┷: " + fileName + " / ??" + rowCount);
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
            log?.Invoke("??욧쉐揶??곕뗄??CSV ???? " + csvPath);
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

        private static bool IsExplicitlyExcludedCategory(Category category)
        {
            if (category == null) return true;

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
            Parameter parameter = FindParameterOnElementOrType(doc, element, parameterName);
            if (parameter == null) return string.Empty;

            try
            {
                string valueString = parameter.AsValueString();
                if (!string.IsNullOrWhiteSpace(valueString)) return valueString;
            }
            catch
            {
            }

            try
            {
                switch (parameter.StorageType)
                {
                    case StorageType.String:
                        return parameter.AsString() ?? string.Empty;
                    case StorageType.Integer:
                        return parameter.AsInteger().ToString();
                    case StorageType.Double:
                        return parameter.AsDouble().ToString(System.Globalization.CultureInfo.InvariantCulture);
                    case StorageType.ElementId:
                        ElementId id = parameter.AsElementId();
                        if (id == null || id == ElementId.InvalidElementId) return string.Empty;
                        Element refElement = doc.GetElement(id);
                        return refElement != null ? (refElement.Name ?? id.IntegerValue.ToString()) : id.IntegerValue.ToString();
                    default:
                        return string.Empty;
                }
            }
            catch
            {
                return string.Empty;
            }
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
            return !string.IsNullOrWhiteSpace(actualName)
                   && string.Equals(actualName.Trim(), expectedName.Trim(), StringComparison.OrdinalIgnoreCase);
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
