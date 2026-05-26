using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;

namespace KKY_Tool_Revit.Services
{
    public static class WorksetAssignmentReviewService
    {
        public const string DefaultExpectedWorksetName = "Workset1";
        private const string ItemLabel = "Workset assignment error check";
        private const string UnknownWorksetName = "Unknown";

        public sealed class Settings
        {
            public string ExpectedWorksetName { get; set; } = DefaultExpectedWorksetName;
            public string FlaggedWorksetName { get; set; } = string.Empty;
            public bool HasAllowedElementScope { get; set; }
            public List<int> AllowedElementIds { get; set; } = new List<int>();
            public List<string> ExtraParameterNames { get; set; } = new List<string>();
        }

        public sealed class ReviewRow
        {
            public string File { get; set; } = string.Empty;
            public string Item { get; set; } = ItemLabel;
            public string Id { get; set; } = string.Empty;
            public string Name { get; set; } = string.Empty;
            public string Result { get; set; } = string.Empty;
            public string Content { get; set; } = string.Empty;
            public string Etc { get; set; } = string.Empty;
            public string Category { get; set; } = string.Empty;
            public string Family { get; set; } = string.Empty;
            public string Comments { get; set; } = string.Empty;
            public Dictionary<string, string> ExtraParams { get; set; } = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        }

        public sealed class FileSummary
        {
            public string File { get; set; } = string.Empty;
            public int TotalReviewed { get; set; }
            public int ErrorCount { get; set; }
            public int OkCount { get; set; }
            public string Status { get; set; } = "pending";
            public string Reason { get; set; } = string.Empty;
        }

        public sealed class ReviewResult
        {
            public string File { get; set; } = string.Empty;
            public string ExpectedWorksetName { get; set; } = DefaultExpectedWorksetName;
            public int TotalReviewed { get; set; }
            public int ErrorCount { get; set; }
            public int OkCount { get; set; }
            public List<ReviewRow> Rows { get; set; } = new List<ReviewRow>();
            public List<FileSummary> FileSummaries { get; set; } = new List<FileSummary>();
        }

        private sealed class TypeInfo
        {
            public string Category { get; set; } = string.Empty;
            public string Family { get; set; } = string.Empty;
            public string TypeName { get; set; } = string.Empty;
        }

        private sealed class ReviewCache
        {
            private readonly Document _doc;
            private readonly WorksetTable _worksetTable;
            private readonly Dictionary<int, string> _worksetNamesById = new Dictionary<int, string>();
            private readonly Dictionary<int, bool> _categoryExclusionById = new Dictionary<int, bool>();
            private readonly Dictionary<int, TypeInfo> _typeInfoById = new Dictionary<int, TypeInfo>();
            private readonly Dictionary<string, string> _extraParameterValues = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            public ReviewCache(Document doc)
            {
                _doc = doc;
                try
                {
                    _worksetTable = doc?.GetWorksetTable();
                }
                catch
                {
                    _worksetTable = null;
                }
            }

            public string GetElementWorksetName(Element element)
            {
                if (element == null)
                {
                    return string.Empty;
                }

                try
                {
                    WorksetId worksetId = element.WorksetId;
                    if (worksetId == null || worksetId == WorksetId.InvalidWorksetId)
                    {
                        return string.Empty;
                    }

                    int key = worksetId.CompatIntegerValue();
                    string cached;
                    if (_worksetNamesById.TryGetValue(key, out cached))
                    {
                        return cached;
                    }

                    Workset workset = _worksetTable?.GetWorkset(worksetId);
                    cached = workset?.Name ?? string.Empty;
                    _worksetNamesById[key] = cached;
                    return cached;
                }
                catch
                {
                    return string.Empty;
                }
            }

            public bool IsExplicitlyExcludedCategory(Category category)
            {
                if (category == null)
                {
                    return true;
                }

                int key = category.Id.CompatIntegerValue();
                bool cached;
                if (_categoryExclusionById.TryGetValue(key, out cached))
                {
                    return cached;
                }

                cached = WorksetAssignmentReviewService.IsExplicitlyExcludedCategory(category);
                _categoryExclusionById[key] = cached;
                return cached;
            }

            public TypeInfo ResolveTypeInfo(Element element)
            {
                int key = element?.Id?.CompatIntegerValue() ?? 0;
                TypeInfo cached;
                if (key > 0 && _typeInfoById.TryGetValue(key, out cached))
                {
                    return cached;
                }

                cached = ResolveTypeInfoCore(_doc, element);
                if (key > 0)
                {
                    _typeInfoById[key] = cached;
                }

                return cached;
            }

            public string GetExtraParameterValue(Element element, string name)
            {
                if (_doc == null || element == null || string.IsNullOrWhiteSpace(name))
                {
                    return string.Empty;
                }

                string normalizedName = name.Trim();
                string key = (element.Id?.CompatIntegerValue() ?? 0).ToString(CultureInfo.InvariantCulture) + "\u001f" + normalizedName;
                string cached;
                if (_extraParameterValues.TryGetValue(key, out cached))
                {
                    return cached;
                }

                cached = ModelParameterExtractionService.GetElementParameterValue(_doc, element, normalizedName) ?? string.Empty;
                _extraParameterValues[key] = cached;
                return cached;
            }
        }

        public static ReviewResult RunOnDocument(Document doc, string fileLabel, Settings settings, Action<double, string> progress = null)
        {
            if (doc == null) throw new ArgumentNullException(nameof(doc));
            if (settings == null) throw new ArgumentNullException(nameof(settings));

            string safeFileLabel = string.IsNullOrWhiteSpace(fileLabel) ? (doc.Title ?? string.Empty) : fileLabel.Trim();
            string expectedWorksetName = NormalizeExpectedWorksetName(settings.ExpectedWorksetName);
            string flaggedWorksetName = NormalizeOptionalWorksetName(settings.FlaggedWorksetName);
            string expectedWorksetKey = NormalizeWorksetKey(expectedWorksetName);
            string flaggedWorksetKey = NormalizeWorksetKey(flaggedWorksetName);
            HashSet<int> allowedElementIds = settings.HasAllowedElementScope
                ? new HashSet<int>((settings.AllowedElementIds ?? Enumerable.Empty<int>()).Where(id => id > 0))
                : null;
            var cache = new ReviewCache(doc);
            List<Element> elements = CollectTargetElements(doc, allowedElementIds, cache);
            List<string> extraParameterNames = NormalizeExtraParameterNames(settings.ExtraParameterNames);

            var result = new ReviewResult
            {
                File = safeFileLabel,
                ExpectedWorksetName = expectedWorksetName,
                TotalReviewed = elements.Count
            };

            int total = Math.Max(elements.Count, 1);
            for (int index = 0; index < elements.Count; index++)
            {
                if (index == 0 || index == elements.Count - 1 || index % 200 == 0)
                {
                    progress?.Invoke(((double)index / total) * 100d, $"Reviewing workset assignment ({index + 1}/{elements.Count})");
                }

                Element element = elements[index];
                string actualWorksetName = cache.GetElementWorksetName(element);
                string actualWorksetKey = NormalizeWorksetKey(actualWorksetName);
                if (!ShouldTreatAsErrorByKey(actualWorksetKey, expectedWorksetKey, flaggedWorksetKey))
                {
                    result.OkCount++;
                    continue;
                }

                result.ErrorCount++;

                TypeInfo typeInfo = cache.ResolveTypeInfo(element);
                string displayWorksetName = string.IsNullOrWhiteSpace(actualWorksetName)
                    ? UnknownWorksetName
                    : actualWorksetName.Trim();

                result.Rows.Add(new ReviewRow
                {
                    File = safeFileLabel,
                    Item = ItemLabel,
                    Id = (element?.Id?.CompatIntegerValue() ?? 0).ToString(CultureInfo.InvariantCulture),
                    Name = typeInfo.TypeName,
                    Result = "Error",
                    Content = string.IsNullOrWhiteSpace(flaggedWorksetName)
                        ? $"[{displayWorksetName}] : The elements are incorrectly assigned to [{displayWorksetName}] workset."
                        : $"[{displayWorksetName}] : The elements are incorrectly assigned to [{flaggedWorksetName}] workset.",
                    Etc = string.Empty,
                    Category = typeInfo.Category,
                    Family = typeInfo.Family,
                    Comments = $"Change the Workset to [{expectedWorksetName}], or consult with the BIM Manager if another workset must be used.",
                    ExtraParams = ReadExtraParamValues(element, extraParameterNames, cache)
                });
            }

            if (result.ErrorCount == 0)
            {
                string okContent = result.TotalReviewed <= 0
                    ? "No target elements found"
                    : $"The {result.TotalReviewed.ToString(CultureInfo.InvariantCulture)} elements are correctly assigned to [{expectedWorksetName}] workset.";

                result.Rows.Add(new ReviewRow
                {
                    File = safeFileLabel,
                    Item = ItemLabel,
                    Id = string.Empty,
                    Name = expectedWorksetName,
                    Result = "OK",
                    Content = okContent,
                    Etc = string.Empty,
                    Category = string.Empty,
                    Family = string.Empty,
                    Comments = string.Empty,
                    ExtraParams = BuildEmptyExtraParamValues(extraParameterNames)
                });
            }

            progress?.Invoke(100d, "Workset assignment review complete");

            result.FileSummaries.Add(new FileSummary
            {
                File = safeFileLabel,
                TotalReviewed = result.TotalReviewed,
                ErrorCount = result.ErrorCount,
                OkCount = result.OkCount,
                Status = "success",
                Reason = BuildSummaryReason(result.TotalReviewed, result.ErrorCount, expectedWorksetName, flaggedWorksetName)
            });

            return result;
        }

        public static DataTable BuildExportTable(IEnumerable<ReviewRow> rows)
        {
            var table = new DataTable("WorksetAssignmentReview");
            table.Columns.Add("Item");
            table.Columns.Add("ID");
            table.Columns.Add("Name");
            table.Columns.Add("Result");
            table.Columns.Add("Content");
            table.Columns.Add("Etc");
            table.Columns.Add("Category");
            table.Columns.Add("Family");
            table.Columns.Add("Comments");

            List<ReviewRow> source = (rows ?? Enumerable.Empty<ReviewRow>())
                .Where(row => row != null)
                .ToList();
            List<string> extraColumns = CollectExtraParamColumns(source);
            foreach (string extraColumn in extraColumns)
            {
                table.Columns.Add(extraColumn);
            }

            foreach (ReviewRow row in source)
            {
                DataRow dataRow = table.NewRow();
                dataRow["Item"] = row.Item ?? ItemLabel;
                dataRow["ID"] = row.Id ?? string.Empty;
                dataRow["Name"] = row.Name ?? string.Empty;
                dataRow["Result"] = row.Result ?? string.Empty;
                dataRow["Content"] = row.Content ?? string.Empty;
                dataRow["Etc"] = row.Etc ?? string.Empty;
                dataRow["Category"] = row.Category ?? string.Empty;
                dataRow["Family"] = row.Family ?? string.Empty;
                dataRow["Comments"] = row.Comments ?? string.Empty;
                foreach (string extraColumn in extraColumns)
                {
                    string value;
                    dataRow[extraColumn] = row.ExtraParams != null && row.ExtraParams.TryGetValue(extraColumn, out value)
                        ? value ?? string.Empty
                        : string.Empty;
                }
                table.Rows.Add(dataRow);
            }

            return table;
        }

        private static string BuildSummaryReason(int totalReviewed, int errorCount, string expectedWorksetName, string flaggedWorksetName)
        {
            if (totalReviewed <= 0)
            {
                return "No target elements found";
            }

            if (!string.IsNullOrWhiteSpace(flaggedWorksetName))
            {
                if (errorCount <= 0)
                {
                    return $"No elements were found in [{flaggedWorksetName}] workset.";
                }

                return $"{errorCount.ToString(CultureInfo.InvariantCulture)} elements were found in [{flaggedWorksetName}] workset.";
            }

            if (errorCount <= 0)
            {
                return $"All {totalReviewed.ToString(CultureInfo.InvariantCulture)} elements are assigned to [{expectedWorksetName}].";
            }

            return $"{errorCount.ToString(CultureInfo.InvariantCulture)} elements are assigned outside [{expectedWorksetName}].";
        }

        private static string NormalizeExpectedWorksetName(string expectedWorksetName)
        {
            string normalized = (expectedWorksetName ?? string.Empty).Trim();
            return string.IsNullOrWhiteSpace(normalized) ? DefaultExpectedWorksetName : normalized;
        }

        private static string NormalizeOptionalWorksetName(string worksetName)
        {
            return (worksetName ?? string.Empty).Trim();
        }

        private static bool ShouldTreatAsError(string actualWorksetName, string expectedWorksetName, string flaggedWorksetName)
        {
            return ShouldTreatAsErrorByKey(
                NormalizeWorksetKey(actualWorksetName),
                NormalizeWorksetKey(expectedWorksetName),
                NormalizeWorksetKey(flaggedWorksetName));
        }

        private static bool ShouldTreatAsErrorByKey(string actualWorksetKey, string expectedWorksetKey, string flaggedWorksetKey)
        {
            if (!string.IsNullOrWhiteSpace(flaggedWorksetKey))
            {
                return IsSameWorksetByKey(actualWorksetKey, flaggedWorksetKey);
            }

            return !IsSameWorksetByKey(actualWorksetKey, expectedWorksetKey);
        }

        private static bool IsExpectedWorkset(string actualWorksetName, string expectedWorksetName)
        {
            return IsSameWorkset(actualWorksetName, expectedWorksetName);
        }

        private static bool IsSameWorkset(string actualWorksetName, string compareWorksetName)
        {
            return IsSameWorksetByKey(NormalizeWorksetKey(actualWorksetName), NormalizeWorksetKey(compareWorksetName));
        }

        private static bool IsSameWorksetByKey(string actualKey, string compareKey)
        {
            if (string.IsNullOrWhiteSpace(actualKey) || string.IsNullOrWhiteSpace(compareKey))
            {
                return false;
            }

            return string.Equals(actualKey, compareKey, StringComparison.OrdinalIgnoreCase);
        }

        private static string NormalizeWorksetKey(string worksetName)
        {
            if (string.IsNullOrWhiteSpace(worksetName))
            {
                return string.Empty;
            }

            return new string((worksetName ?? string.Empty)
                .Trim()
                .Where(ch => !char.IsWhiteSpace(ch))
                .ToArray())
                .ToLowerInvariant();
        }

        private static string GetElementWorksetName(Document doc, Element element)
        {
            if (doc == null || element == null)
            {
                return string.Empty;
            }

            try
            {
                WorksetId worksetId = element.WorksetId;
                if (worksetId == null || worksetId == WorksetId.InvalidWorksetId)
                {
                    return string.Empty;
                }

                WorksetTable table = doc.GetWorksetTable();
                if (table == null)
                {
                    return string.Empty;
                }

                Workset workset = table.GetWorkset(worksetId);
                return workset?.Name ?? string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }

        private static TypeInfo ResolveTypeInfoCore(Document doc, Element element)
        {
            var info = new TypeInfo
            {
                Category = element?.Category?.Name ?? string.Empty,
                TypeName = element?.Name ?? string.Empty
            };

            if (element is FamilyInstance familyInstance)
            {
                FamilySymbol symbol = familyInstance.Symbol;
                if (symbol != null)
                {
                    info.Family = symbol.FamilyName ?? string.Empty;
                    if (string.IsNullOrWhiteSpace(info.TypeName))
                    {
                        info.TypeName = symbol.Name ?? string.Empty;
                    }
                }
            }

            ElementType elementType = null;
            try
            {
                ElementId typeId = element?.GetTypeId();
                if (typeId != null && typeId != ElementId.InvalidElementId)
                {
                    elementType = doc.GetElement(typeId) as ElementType;
                }
            }
            catch
            {
                elementType = null;
            }

            if (elementType != null)
            {
                if (string.IsNullOrWhiteSpace(info.Family))
                {
                    info.Family = elementType.FamilyName ?? string.Empty;
                }

                if (!string.IsNullOrWhiteSpace(elementType.Name))
                {
                    info.TypeName = elementType.Name;
                }
            }

            if (string.IsNullOrWhiteSpace(info.TypeName))
            {
                info.TypeName = element?.Name ?? string.Empty;
            }

            return info;
        }

        private static List<Element> CollectTargetElements(Document doc, HashSet<int> allowedElementIds, ReviewCache cache)
        {
            if (allowedElementIds != null)
            {
                if (allowedElementIds.Count == 0)
                {
                    return new List<Element>();
                }

                return allowedElementIds
                    .OrderBy(id => id)
                    .Select(id => TryGetElementById(doc, id))
                    .Where(element => ShouldReviewElement(element, cache))
                    .ToList();
            }

            return new FilteredElementCollector(doc)
                .WhereElementIsNotElementType()
                .Cast<Element>()
                .Where(element => ShouldReviewElement(element, cache))
                .ToList();
        }

        private static Element TryGetElementById(Document doc, int elementId)
        {
            if (doc == null || elementId <= 0)
            {
                return null;
            }

            try
            {
#pragma warning disable CS0618 // Revit 2019-2023 compatibility uses the int ElementId constructor.
                return doc.GetElement(new ElementId(elementId));
#pragma warning restore CS0618
            }
            catch
            {
                return null;
            }
        }

        private static bool ShouldReviewElement(Element element, ReviewCache cache)
        {
            if (element == null) return false;
            if (element.ViewSpecific) return false;
            if (element.Category == null) return false;
            if (element.Category.CategoryType != CategoryType.Model) return false;

            string categoryName = element.Category.Name ?? string.Empty;
            if (element is Level) return false;
            if (element is ReferencePlane) return false;
            if (element is CurveElement) return false;
            if (element is Grid) return false;
            if (element is Group) return false;
            if (element is AssemblyInstance) return false;
            if (element is RevitLinkInstance) return false;
            if (element is ImportInstance) return false;
            if (element is View) return false;
            if (element is ElementType) return false;
            if (element is BasePoint) return false;
            if (element is Room) return false;
            if (element is Area) return false;
            if (element is MEPSystem) return false;
            if (IsExcludedExternalReferenceElement(element)) return false;
            if (string.IsNullOrWhiteSpace(categoryName)) return false;

            int categoryId = element.Category.Id.CompatIntegerValue();
            if (categoryId == (int)BuiltInCategory.OST_Levels) return false;
            if (categoryId == (int)BuiltInCategory.OST_Grids) return false;
            if (categoryId == (int)BuiltInCategory.OST_PointClouds) return false;
            if (categoryId == (int)BuiltInCategory.OST_RvtLinks) return false;
            if (categoryId == (int)BuiltInCategory.OST_Cameras) return false;
            if (categoryId == (int)BuiltInCategory.OST_SectionBox) return false;
            if (categoryId == (int)BuiltInCategory.OST_VolumeOfInterest) return false;
            return !(cache?.IsExplicitlyExcludedCategory(element.Category) ?? IsExplicitlyExcludedCategory(element.Category));
        }

        private static bool IsExcludedExternalReferenceElement(Element element)
        {
            if (element == null) return false;
            return IsElementOfType(element, "Autodesk.Revit.DB.PointClouds.PointCloudInstance");
        }

        private static bool IsElementOfType(Element element, string fullTypeName)
        {
            if (element == null || string.IsNullOrWhiteSpace(fullTypeName)) return false;

            Type currentType = element.GetType();
            while (currentType != null)
            {
                if (string.Equals(currentType.FullName, fullTypeName, StringComparison.Ordinal))
                {
                    return true;
                }

                currentType = currentType.BaseType;
            }

            return false;
        }

        private static bool IsExplicitlyExcludedCategory(Category category)
        {
            if (category == null) return true;

            string normalized = (category.Name ?? string.Empty).Trim();
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
                "Material",
                "Project Information",
                "Sun Path",
                "Pipe Segment",
                "Primary Contour",
                "Legend Component",
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
                "OST_Materials",
                "OST_IOSDatumPlane",
                "OST_VolumeOfInterest",
                "OST_SectionBox");
        }

        private static List<string> NormalizeExtraParameterNames(IEnumerable<string> extraParameterNames)
        {
            var result = new List<string>();
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (string name in extraParameterNames ?? Enumerable.Empty<string>())
            {
                string normalized = (name ?? string.Empty).Trim();
                if (string.IsNullOrWhiteSpace(normalized)) continue;
                if (seen.Add(normalized))
                {
                    result.Add(normalized);
                }
            }

            return result;
        }

        private static Dictionary<string, string> ReadExtraParamValues(Element element, IEnumerable<string> extraParameterNames, ReviewCache cache)
        {
            var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            if (cache == null || element == null) return result;

            foreach (string name in extraParameterNames ?? Enumerable.Empty<string>())
            {
                if (string.IsNullOrWhiteSpace(name)) continue;
                result[name] = cache.GetExtraParameterValue(element, name);
            }

            return result;
        }

        private static Dictionary<string, string> BuildEmptyExtraParamValues(IEnumerable<string> extraParameterNames)
        {
            var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (string name in extraParameterNames ?? Enumerable.Empty<string>())
            {
                if (string.IsNullOrWhiteSpace(name)) continue;
                result[name] = string.Empty;
            }

            return result;
        }

        private static List<string> CollectExtraParamColumns(IEnumerable<ReviewRow> rows)
        {
            var result = new List<string>();
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (ReviewRow row in rows ?? Enumerable.Empty<ReviewRow>())
            {
                if (row?.ExtraParams == null) continue;
                foreach (string key in row.ExtraParams.Keys)
                {
                    string normalized = (key ?? string.Empty).Trim();
                    if (string.IsNullOrWhiteSpace(normalized)) continue;
                    if (seen.Add(normalized))
                    {
                        result.Add(normalized);
                    }
                }
            }

            return result;
        }

        private static bool MatchesBuiltInCategoryNames(Category category, params string[] builtInCategoryNames)
        {
            if (category == null || builtInCategoryNames == null || builtInCategoryNames.Length == 0)
            {
                return false;
            }

            string actualName;
            try
            {
                actualName = Enum.GetName(typeof(BuiltInCategory), category.Id.CompatIntegerValue()) ?? string.Empty;
            }
            catch
            {
                return false;
            }

            if (string.IsNullOrWhiteSpace(actualName))
            {
                return false;
            }

            return builtInCategoryNames.Any(name => string.Equals(actualName, name, StringComparison.OrdinalIgnoreCase));
        }
    }
}
