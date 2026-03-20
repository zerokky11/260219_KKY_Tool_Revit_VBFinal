using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;

namespace KKY_Tool_Revit.Services
{
    public static class FloorInfoReviewService
    {
        private const double FootToMillimeter = 304.8d;
        private const double BoundaryEpsilonFt = 1.0d / 304.8d;

        public sealed class Settings
        {
            public string ParameterName { get; set; } = string.Empty;
            public List<LevelRule> LevelRules { get; set; } = new List<LevelRule>();
        }

        public sealed class LevelRule
        {
            public string LevelName { get; set; } = string.Empty;
            public double AbsoluteZFt { get; set; }
            public bool UseAsBoundary { get; set; } = true;
            public string ExpectedValue { get; set; } = string.Empty;
        }

        public sealed class LevelOption
        {
            public int LevelId { get; set; }
            public string LevelName { get; set; } = string.Empty;
            public double AbsoluteZFt { get; set; }
            public double AbsoluteZMm { get; set; }
        }

        public sealed class ConfigSnapshot
        {
            public string DocumentTitle { get; set; } = string.Empty;
            public List<LevelOption> Levels { get; set; } = new List<LevelOption>();
            public List<string> Warnings { get; set; } = new List<string>();
        }

        public sealed class ReviewRow
        {
            public string File { get; set; } = string.Empty;
            public int ElementId { get; set; }
            public string Category { get; set; } = string.Empty;
            public string Family { get; set; } = string.Empty;
            public string TypeName { get; set; } = string.Empty;
            public string ElementName { get; set; } = string.Empty;
            public string ParameterName { get; set; } = string.Empty;
            public string LowerLevelName { get; set; } = string.Empty;
            public string UpperLevelName { get; set; } = string.Empty;
            public string ExpectedValue { get; set; } = string.Empty;
            public string ActualValue { get; set; } = string.Empty;
            public double RepresentativeZFt { get; set; }
            public double RepresentativeZMm { get; set; }
            public double BottomZFt { get; set; }
            public double BottomZMm { get; set; }
            public double TopZFt { get; set; }
            public double TopZMm { get; set; }
            public bool SpansMultipleZones { get; set; }
            public string Result { get; set; } = string.Empty;
            public string Note { get; set; } = string.Empty;
        }

        public sealed class FileSummary
        {
            public string File { get; set; } = string.Empty;
            public int Total { get; set; }
            public int Issues { get; set; }
            public int Near { get; set; }
            public string Status { get; set; } = "pending";
            public string Reason { get; set; } = string.Empty;
        }

        public sealed class ReviewResult
        {
            public string File { get; set; } = string.Empty;
            public string ParameterName { get; set; } = string.Empty;
            public int TotalElements { get; set; }
            public int EvaluatedElements { get; set; }
            public int IssueCount { get; set; }
            public int MismatchCount { get; set; }
            public int MissingParameterCount { get; set; }
            public int MissingRuleCount { get; set; }
            public int MissingGeometryCount { get; set; }
            public List<ReviewRow> Rows { get; set; } = new List<ReviewRow>();
            public List<FileSummary> FileSummaries { get; set; } = new List<FileSummary>();
            public List<string> Warnings { get; set; } = new List<string>();
        }

        public static ConfigSnapshot ReadConfig(Document doc)
        {
            if (doc == null) throw new ArgumentNullException(nameof(doc));

            List<Level> levels = CollectLevels(doc);
            if (levels.Count == 0)
            {
                throw new InvalidOperationException("현재 문서에 레벨이 없습니다.");
            }

            var snapshot = new ConfigSnapshot
            {
                DocumentTitle = string.IsNullOrWhiteSpace(doc.Title) ? string.Empty : doc.Title
            };

            foreach (Level level in levels)
            {
                double zFt = GetAbsoluteLevelZ(level);
                snapshot.Levels.Add(new LevelOption
                {
                    LevelId = level.Id.IntegerValue,
                    LevelName = level.Name ?? string.Empty,
                    AbsoluteZFt = Round(zFt, 6),
                    AbsoluteZMm = Round(ToMillimeters(zFt), 1)
                });
            }

            snapshot.Warnings.AddRange(BuildDuplicateLevelWarnings(snapshot.Levels));
            return snapshot;
        }

        public static ReviewResult RunOnDocument(Document doc, string fileLabel, Settings settings, Action<double, string> progress = null)
        {
            if (doc == null) throw new ArgumentNullException(nameof(doc));
            if (settings == null) throw new ArgumentNullException(nameof(settings));

            string safeFileLabel = string.IsNullOrWhiteSpace(fileLabel) ? (doc.Title ?? string.Empty) : fileLabel;
            List<Level> levels = CollectLevels(doc);
            if (levels.Count == 0)
            {
                throw new InvalidOperationException("검토 가능한 레벨이 없습니다.");
            }

            string parameterName = (settings.ParameterName ?? string.Empty).Trim();
            if (string.IsNullOrWhiteSpace(parameterName))
            {
                throw new InvalidOperationException("층정보 검토 대상 파라미터명이 비어 있습니다.");
            }

            var ruleMap = BuildRuleMap(settings.LevelRules, levels);
            List<Level> zoneLevels = BuildZoneLevels(settings.LevelRules, levels);
            if (zoneLevels.Count == 0)
            {
                throw new InvalidOperationException("층정보 영역을 구분할 레벨을 최소 1개 이상 선택해야 합니다.");
            }
            var result = new ReviewResult
            {
                File = safeFileLabel,
                ParameterName = parameterName
            };

            result.Warnings.AddRange(BuildDuplicateLevelWarnings(levels.Select(level => new LevelOption
            {
                LevelId = level.Id.IntegerValue,
                LevelName = level.Name ?? string.Empty,
                AbsoluteZFt = GetAbsoluteLevelZ(level)
            }).ToList()));

            List<Element> elements = CollectTargetElements(doc);
            result.TotalElements = elements.Count;
            int total = Math.Max(elements.Count, 1);

            for (int index = 0; index < elements.Count; index++)
            {
                Element element = elements[index];
                progress?.Invoke(((double)index / total) * 100d, $"층정보 검토 중 ({index + 1}/{elements.Count})");

                BoundingBoxXYZ bbox = SafeGetBoundingBox(element);
                if (bbox == null)
                {
                    result.MissingGeometryCount++;
                    result.IssueCount++;
                    result.Rows.Add(BuildIssueRow(safeFileLabel, element, parameterName, null, null, null, string.Empty, string.Empty, 0d, 0d, 0d, false, "NO_GEOMETRY", "BoundingBox를 가져오지 못했습니다."));
                    continue;
                }

                result.EvaluatedElements++;

                double bottomZ = Math.Min(bbox.Min.Z, bbox.Max.Z);
                double topZ = Math.Max(bbox.Min.Z, bbox.Max.Z);
                int bottomZone = FindZoneIndex(zoneLevels, bottomZ);
                int topZone = FindZoneIndex(zoneLevels, Math.Max(bottomZ, topZ - BoundaryEpsilonFt));
                bool spansMultipleZones = bottomZone != topZone;
                double representativeZ = ResolveRepresentativeZ(element, bbox, spansMultipleZones);
                int reviewZone = spansMultipleZones ? bottomZone : FindZoneIndex(zoneLevels, representativeZ);

                Level lowerLevel = zoneLevels[Math.Max(0, Math.Min(reviewZone, zoneLevels.Count - 1))];
                Level upperLevel = reviewZone + 1 < zoneLevels.Count ? zoneLevels[reviewZone + 1] : null;

                string expectedValue = string.Empty;
                if (!ruleMap.TryGetValue(NormalizeKey(lowerLevel.Name), out expectedValue))
                {
                    result.MissingRuleCount++;
                    result.IssueCount++;
                    result.Rows.Add(BuildIssueRow(
                        safeFileLabel,
                        element,
                        parameterName,
                        lowerLevel,
                        upperLevel,
                        bbox,
                        string.Empty,
                        string.Empty,
                        representativeZ,
                        bottomZ,
                        topZ,
                        spansMultipleZones,
                        "RULE_MISSING",
                        $"레벨 '{lowerLevel.Name}' 에 대한 기대 층정보 값이 설정되지 않았습니다."));
                    continue;
                }

                ParameterMatch parameterMatch = FindParameterValue(element, parameterName);
                if (!parameterMatch.Found)
                {
                    result.MissingParameterCount++;
                    result.IssueCount++;
                    result.Rows.Add(BuildIssueRow(
                        safeFileLabel,
                        element,
                        parameterName,
                        lowerLevel,
                        upperLevel,
                        bbox,
                        expectedValue,
                        string.Empty,
                        representativeZ,
                        bottomZ,
                        topZ,
                        spansMultipleZones,
                        "PARAMETER_MISSING",
                        "대상 파라미터를 찾지 못했습니다."));
                    continue;
                }

                string actualValue = NormalizeValue(parameterMatch.Value);
                string expectedNormalized = NormalizeValue(expectedValue);
                if (!string.Equals(actualValue, expectedNormalized, StringComparison.OrdinalIgnoreCase))
                {
                    result.MismatchCount++;
                    result.IssueCount++;
                    result.Rows.Add(BuildIssueRow(
                        safeFileLabel,
                        element,
                        parameterName,
                        lowerLevel,
                        upperLevel,
                        bbox,
                        expectedValue,
                        parameterMatch.Value,
                        representativeZ,
                        bottomZ,
                        topZ,
                        spansMultipleZones,
                        "MISMATCH",
                        spansMultipleZones
                            ? "여러 레벨 구간을 관통하여 가장 아래 구간의 층정보를 기대값으로 사용했습니다."
                            : "객체의 층정보 값이 기대값과 일치하지 않습니다."));
                }
            }

            progress?.Invoke(100d, "층정보 검토 완료");
            result.FileSummaries.Add(new FileSummary
            {
                File = safeFileLabel,
                Total = result.EvaluatedElements,
                Issues = result.IssueCount,
                Near = 0,
                Status = "success",
                Reason = string.Empty
            });
            return result;
        }

        public static DataTable BuildExportTable(IEnumerable<ReviewRow> rows)
        {
            var table = new DataTable("FloorInfoReview");
            table.Columns.Add("File");
            table.Columns.Add("ElementId");
            table.Columns.Add("Category");
            table.Columns.Add("Family");
            table.Columns.Add("TypeName");
            table.Columns.Add("ElementName");
            table.Columns.Add("ParameterName");
            table.Columns.Add("LowerLevel");
            table.Columns.Add("UpperLevel");
            table.Columns.Add("ExpectedValue");
            table.Columns.Add("ActualValue");
            table.Columns.Add("RepresentativeZ (mm)", typeof(double));
            table.Columns.Add("BottomZ (mm)", typeof(double));
            table.Columns.Add("TopZ (mm)", typeof(double));
            table.Columns.Add("SpansMultipleZones");
            table.Columns.Add("Result");
            table.Columns.Add("Note");

            List<ReviewRow> source = (rows ?? Enumerable.Empty<ReviewRow>()).Where(row => row != null).ToList();
            if (source.Count == 0)
            {
                DataRow empty = table.NewRow();
                empty["File"] = "오류가 없습니다.";
                table.Rows.Add(empty);
                return table;
            }

            foreach (ReviewRow row in source)
            {
                DataRow dataRow = table.NewRow();
                dataRow["File"] = row.File ?? string.Empty;
                dataRow["ElementId"] = row.ElementId.ToString(CultureInfo.InvariantCulture);
                dataRow["Category"] = row.Category ?? string.Empty;
                dataRow["Family"] = row.Family ?? string.Empty;
                dataRow["TypeName"] = row.TypeName ?? string.Empty;
                dataRow["ElementName"] = row.ElementName ?? string.Empty;
                dataRow["ParameterName"] = row.ParameterName ?? string.Empty;
                dataRow["LowerLevel"] = row.LowerLevelName ?? string.Empty;
                dataRow["UpperLevel"] = row.UpperLevelName ?? string.Empty;
                dataRow["ExpectedValue"] = row.ExpectedValue ?? string.Empty;
                dataRow["ActualValue"] = row.ActualValue ?? string.Empty;
                dataRow["RepresentativeZ (mm)"] = Round(row.RepresentativeZMm, 1);
                dataRow["BottomZ (mm)"] = Round(row.BottomZMm, 1);
                dataRow["TopZ (mm)"] = Round(row.TopZMm, 1);
                dataRow["SpansMultipleZones"] = row.SpansMultipleZones ? "Y" : "N";
                dataRow["Result"] = row.Result ?? string.Empty;
                dataRow["Note"] = row.Note ?? string.Empty;
                table.Rows.Add(dataRow);
            }

            return table;
        }

        private static List<Level> CollectLevels(Document doc)
        {
            return new FilteredElementCollector(doc)
                .OfClass(typeof(Level))
                .Cast<Level>()
                .Where(level => level != null)
                .OrderBy(level => GetAbsoluteLevelZ(level))
                .ThenBy(level => level.Name ?? string.Empty, StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        private static List<Element> CollectTargetElements(Document doc)
        {
            return new FilteredElementCollector(doc)
                .WhereElementIsNotElementType()
                .Cast<Element>()
                .Where(ShouldReviewElement)
                .ToList();
        }

        private static bool ShouldReviewElement(Element element)
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
            if (string.IsNullOrWhiteSpace(categoryName)) return false;

            int categoryId = element.Category.Id.IntegerValue;
            if (categoryId == (int)BuiltInCategory.OST_Levels) return false;
            if (categoryId == (int)BuiltInCategory.OST_Grids) return false;
            if (categoryId == (int)BuiltInCategory.OST_RvtLinks) return false;
            if (categoryId == (int)BuiltInCategory.OST_Cameras) return false;
            if (categoryId == (int)BuiltInCategory.OST_SectionBox) return false;
            if (categoryId == (int)BuiltInCategory.OST_VolumeOfInterest) return false;
            return !IsExplicitlyExcludedCategory(element.Category);
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
                "OST_Materials",
                "OST_IOSDatumPlane",
                "OST_VolumeOfInterest",
                "OST_SectionBox");
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
                actualName = Enum.GetName(typeof(BuiltInCategory), category.Id.IntegerValue) ?? string.Empty;
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

        private static Dictionary<string, string> BuildRuleMap(IEnumerable<LevelRule> rules, IList<Level> levels)
        {
            var map = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            foreach (LevelRule rule in rules ?? Enumerable.Empty<LevelRule>())
            {
                if (rule == null) continue;
                string levelName = (rule.LevelName ?? string.Empty).Trim();
                if (string.IsNullOrWhiteSpace(levelName)) continue;
                map[NormalizeKey(levelName)] = (rule.ExpectedValue ?? string.Empty).Trim();
            }

            foreach (Level level in levels ?? Enumerable.Empty<Level>())
            {
                if (level == null) continue;
                string key = NormalizeKey(level.Name);
                if (!map.ContainsKey(key))
                {
                    map[key] = string.Empty;
                }
            }

            return map;
        }

        private static List<Level> BuildZoneLevels(IEnumerable<LevelRule> rules, IList<Level> levels)
        {
            List<Level> source = (levels ?? Array.Empty<Level>())
                .Where(level => level != null)
                .OrderBy(level => GetAbsoluteLevelZ(level))
                .ThenBy(level => level.Name ?? string.Empty, StringComparer.OrdinalIgnoreCase)
                .ToList();

            List<LevelRule> configuredRules = (rules ?? Enumerable.Empty<LevelRule>())
                .Where(rule => rule != null)
                .ToList();

            if (configuredRules.Count == 0)
            {
                return source;
            }

            var selected = new HashSet<string>(
                configuredRules
                    .Where(rule => rule.UseAsBoundary)
                    .Select(rule => NormalizeKey(rule.LevelName))
                    .Where(name => !string.IsNullOrWhiteSpace(name)),
                StringComparer.OrdinalIgnoreCase);

            if (selected.Count == 0)
            {
                return new List<Level>();
            }

            return source
                .Where(level => selected.Contains(NormalizeKey(level.Name)))
                .ToList();
        }

        private static IEnumerable<string> BuildDuplicateLevelWarnings(IList<LevelOption> levels)
        {
            if (levels == null || levels.Count == 0) yield break;

            foreach (IGrouping<string, LevelOption> group in levels.GroupBy(level => NormalizeKey(level.LevelName)))
            {
                if (group.Count() < 2) continue;
                string name = group.First().LevelName ?? string.Empty;
                string positions = string.Join(", ", group.Select(item => Round(item.AbsoluteZMm, 1).ToString("0.0", CultureInfo.InvariantCulture) + "mm"));
                yield return $"동일한 레벨명 '{name}' 이 여러 높이에 존재합니다: {positions}";
            }
        }

        private static double ResolveRepresentativeZ(Element element, BoundingBoxXYZ bbox, bool spansMultipleZones)
        {
            double bottom = Math.Min(bbox.Min.Z, bbox.Max.Z);
            double top = Math.Max(bbox.Min.Z, bbox.Max.Z);

            if (spansMultipleZones)
            {
                return bottom;
            }

            if (element.Location is LocationPoint point && point.Point != null)
            {
                return point.Point.Z;
            }

            if (element.Location is LocationCurve curve && curve.Curve != null)
            {
                try
                {
                    XYZ start = curve.Curve.GetEndPoint(0);
                    XYZ end = curve.Curve.GetEndPoint(1);
                    return (start.Z + end.Z) / 2d;
                }
                catch
                {
                }
            }

            return (bottom + top) / 2d;
        }

        private static ParameterMatch FindParameterValue(Element element, string parameterName)
        {
            Parameter parameter = FindParameter(element, parameterName);
            if (parameter != null)
            {
                return new ParameterMatch(true, ReadParameterValue(parameter));
            }

            ElementType elementType = null;
            try
            {
                ElementId typeId = element.GetTypeId();
                if (typeId != null && typeId != ElementId.InvalidElementId)
                {
                    elementType = element.Document.GetElement(typeId) as ElementType;
                }
            }
            catch
            {
                elementType = null;
            }

            parameter = FindParameter(elementType, parameterName);
            return parameter != null
                ? new ParameterMatch(true, ReadParameterValue(parameter))
                : new ParameterMatch(false, string.Empty);
        }

        private static Parameter FindParameter(Element element, string parameterName)
        {
            if (element == null || string.IsNullOrWhiteSpace(parameterName))
            {
                return null;
            }

            Parameter direct = null;
            try
            {
                direct = element.LookupParameter(parameterName);
            }
            catch
            {
                direct = null;
            }

            if (direct != null)
            {
                return direct;
            }

            foreach (Parameter parameter in element.Parameters.Cast<Parameter>())
            {
                if (parameter?.Definition == null) continue;
                string definitionName = parameter.Definition.Name ?? string.Empty;
                if (string.Equals(definitionName.Trim(), parameterName.Trim(), StringComparison.OrdinalIgnoreCase))
                {
                    return parameter;
                }
            }

            return null;
        }

        private static string ReadParameterValue(Parameter parameter)
        {
            if (parameter == null) return string.Empty;

            try
            {
                switch (parameter.StorageType)
                {
                    case StorageType.String:
                        return parameter.AsString() ?? string.Empty;
                    case StorageType.Integer:
                        return parameter.AsInteger().ToString(CultureInfo.InvariantCulture);
                    case StorageType.Double:
                        return parameter.AsValueString() ?? parameter.AsDouble().ToString(CultureInfo.InvariantCulture);
                    case StorageType.ElementId:
                        ElementId id = parameter.AsElementId();
                        return id == null ? string.Empty : id.IntegerValue.ToString(CultureInfo.InvariantCulture);
                    default:
                        return parameter.AsValueString() ?? string.Empty;
                }
            }
            catch
            {
                return string.Empty;
            }
        }

        private static BoundingBoxXYZ SafeGetBoundingBox(Element element)
        {
            if (element == null) return null;

            try
            {
                return element.get_BoundingBox(null);
            }
            catch
            {
                return null;
            }
        }

        private static int FindZoneIndex(IList<Level> levels, double z)
        {
            if (levels == null || levels.Count == 0) return 0;

            for (int index = levels.Count - 1; index >= 0; index--)
            {
                double levelZ = GetAbsoluteLevelZ(levels[index]);
                if (z >= levelZ - BoundaryEpsilonFt)
                {
                    return index;
                }
            }

            return 0;
        }

        private static string NormalizeKey(string value)
        {
            return (value ?? string.Empty).Trim();
        }

        private static string NormalizeValue(string value)
        {
            return (value ?? string.Empty).Trim();
        }

        private static double GetAbsoluteLevelZ(Level level)
        {
            if (level == null) return 0d;
            try
            {
                return level.Elevation;
            }
            catch
            {
                return 0d;
            }
        }

        private static double ToMillimeters(double feet)
        {
            return feet * FootToMillimeter;
        }

        private static double Round(double value, int digits)
        {
            return Math.Round(value, digits, MidpointRounding.AwayFromZero);
        }

        private static ReviewRow BuildIssueRow(
            string file,
            Element element,
            string parameterName,
            Level lowerLevel,
            Level upperLevel,
            BoundingBoxXYZ bbox,
            string expectedValue,
            string actualValue,
            double representativeZ,
            double bottomZ,
            double topZ,
            bool spansMultipleZones,
            string result,
            string note)
        {
            string familyName = string.Empty;
            string typeName = string.Empty;

            if (element is FamilyInstance familyInstance)
            {
                familyName = familyInstance.Symbol?.FamilyName ?? string.Empty;
                typeName = familyInstance.Symbol?.Name ?? string.Empty;
            }
            else
            {
                ElementType elementType = null;
                try
                {
                    ElementId typeId = element.GetTypeId();
                    if (typeId != null && typeId != ElementId.InvalidElementId)
                    {
                        elementType = element.Document.GetElement(typeId) as ElementType;
                    }
                }
                catch
                {
                    elementType = null;
                }

                typeName = elementType?.Name ?? string.Empty;
                familyName = elementType?.FamilyName ?? string.Empty;
            }

            return new ReviewRow
            {
                File = file ?? string.Empty,
                ElementId = element?.Id?.IntegerValue ?? 0,
                Category = element?.Category?.Name ?? string.Empty,
                Family = familyName,
                TypeName = typeName,
                ElementName = element?.Name ?? string.Empty,
                ParameterName = parameterName ?? string.Empty,
                LowerLevelName = lowerLevel?.Name ?? string.Empty,
                UpperLevelName = upperLevel?.Name ?? string.Empty,
                ExpectedValue = expectedValue ?? string.Empty,
                ActualValue = actualValue ?? string.Empty,
                RepresentativeZFt = Round(representativeZ, 6),
                RepresentativeZMm = Round(ToMillimeters(representativeZ), 1),
                BottomZFt = Round(bottomZ, 6),
                BottomZMm = Round(ToMillimeters(bottomZ), 1),
                TopZFt = Round(topZ, 6),
                TopZMm = Round(ToMillimeters(topZ), 1),
                SpansMultipleZones = spansMultipleZones,
                Result = result ?? string.Empty,
                Note = note ?? string.Empty
            };
        }

        private readonly struct ParameterMatch
        {
            public ParameterMatch(bool found, string value)
            {
                Found = found;
                Value = value ?? string.Empty;
            }

            public bool Found { get; }
            public string Value { get; }
        }
    }
}
