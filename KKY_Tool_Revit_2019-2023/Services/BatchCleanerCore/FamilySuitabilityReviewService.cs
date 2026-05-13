using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.DB.Mechanical;

namespace KKY_Tool_Revit.Services
{
    public static class FamilySuitabilityReviewService
    {
        private static readonly Regex MultiWhitespace = new Regex(@"\s+", RegexOptions.Compiled);

        public sealed class Settings
        {
            public string MatchReviewText { get; set; } = string.Empty;
            public string MismatchReviewText { get; set; } = string.Empty;
            public List<CriteriaRule> CriteriaRules { get; set; } = new List<CriteriaRule>();
            public List<FilterRule> FilterRules { get; set; } = new List<FilterRule>();
        }

        public sealed class CriteriaRule
        {
            public string Category { get; set; } = string.Empty;
            public string Family { get; set; } = string.Empty;
            public string TypeName { get; set; } = string.Empty;
        }

        public sealed class FilterRule
        {
            public string Target { get; set; } = "familyOrType";
            public string Keyword { get; set; } = string.Empty;
            public string ReviewText { get; set; } = string.Empty;
        }

        public sealed class ReviewRow
        {
            public string File { get; set; } = string.Empty;
            public string Category { get; set; } = string.Empty;
            public string Family { get; set; } = string.Empty;
            public string TypeName { get; set; } = string.Empty;
            public int ElementCount { get; set; }
            public string Review { get; set; } = string.Empty;
            public string ReviewSource { get; set; } = string.Empty;
            public string Status { get; set; } = string.Empty;
        }

        public sealed class FileSummary
        {
            public string File { get; set; } = string.Empty;
            public int Total { get; set; }
            public int Issues { get; set; }
            public int Near { get; set; }
            public int MatchCount { get; set; }
            public int MismatchCount { get; set; }
            public int FilterCount { get; set; }
            public int ElementCount { get; set; }
            public string Status { get; set; } = "pending";
            public string Reason { get; set; } = string.Empty;
        }

        public sealed class ReviewResult
        {
            public string File { get; set; } = string.Empty;
            public int CriteriaCount { get; set; }
            public int TotalElements { get; set; }
            public int TotalGroups { get; set; }
            public int MatchCount { get; set; }
            public int MismatchCount { get; set; }
            public int FilterCount { get; set; }
            public List<ReviewRow> Rows { get; set; } = new List<ReviewRow>();
            public List<FileSummary> FileSummaries { get; set; } = new List<FileSummary>();
            public List<string> Warnings { get; set; } = new List<string>();
        }

        private sealed class UsageGroup
        {
            public string Key { get; set; } = string.Empty;
            public string Category { get; set; } = string.Empty;
            public string Family { get; set; } = string.Empty;
            public string TypeName { get; set; } = string.Empty;
            public string NormalizedFamily { get; set; } = string.Empty;
            public string NormalizedTypeName { get; set; } = string.Empty;
            public int ElementCount { get; set; }
        }

        public static ReviewResult RunOnDocument(Document doc, string fileLabel, Settings settings, Action<double, string> progress = null)
        {
            if (doc == null) throw new ArgumentNullException(nameof(doc));
            if (settings == null) throw new ArgumentNullException(nameof(settings));

            string safeFileLabel = string.IsNullOrWhiteSpace(fileLabel) ? (doc.Title ?? string.Empty) : fileLabel.Trim();
#if false
            string matchReviewText = string.IsNullOrWhiteSpace(settings.MatchReviewText) ? "기준 일치" : settings.MatchReviewText.Trim();
            string mismatchReviewText = string.IsNullOrWhiteSpace(settings.MismatchReviewText) ? "기준 미일치" : settings.MismatchReviewText.Trim();
 #endif
            string matchReviewText = string.IsNullOrWhiteSpace(settings.MatchReviewText) ? "\uAE30\uC900 \uC77C\uCE58" : settings.MatchReviewText.Trim();
            string mismatchReviewText = string.IsNullOrWhiteSpace(settings.MismatchReviewText) ? "\uAE30\uC900 \uBBF8\uC77C\uCE58" : settings.MismatchReviewText.Trim();
            HashSet<string> criteriaSet = BuildCriteriaSet(settings.CriteriaRules);
#if false
            if (criteriaSet.Count == 0)
            {
                throw new InvalidOperationException("기준 엑셀에서 사용할 Category / Family / Type 조합이 없습니다.");
            }

 #endif
            if (criteriaSet.Count == 0)
            {
                throw new InvalidOperationException("\uAE30\uC900 \uC124\uC815\uC5D0\uC11C \uC0AC\uC6A9\uD560 Category / Family / Type \uC870\uD569\uC774 \uC5C6\uC2B5\uB2C8\uB2E4.");
            }

            List<FilterRule> filterRules = NormalizeFilterRules(settings.FilterRules);
            List<Element> elements = CollectTargetElements(doc);
            var usageMap = new Dictionary<string, UsageGroup>(StringComparer.OrdinalIgnoreCase);
            int totalElements = Math.Max(elements.Count, 1);

            for (int index = 0; index < elements.Count; index++)
            {
                if (index == 0 || index == elements.Count - 1 || index % 200 == 0)
                {
#if false
                    progress?.Invoke(((double)index / totalElements) * 75d, $"사용 객체 집계 중 ({index + 1}/{elements.Count})");
                }

 #endif
                    progress?.Invoke(((double)index / totalElements) * 75d, $"\uC0AC\uC6A9 \uAC1D\uCCB4 \uC9D1\uACC4 \uC911 ({index + 1}/{elements.Count})");
                }
                UsageGroup usage = TryBuildUsageGroup(doc, elements[index]);
                if (usage == null) continue;

                UsageGroup existing;
                if (usageMap.TryGetValue(usage.Key, out existing))
                {
                    existing.ElementCount += 1;
                }
                else
                {
                    usage.ElementCount = 1;
                    usageMap[usage.Key] = usage;
                }
            }

            List<UsageGroup> groups = usageMap.Values
                .OrderBy(item => item.Category ?? string.Empty, StringComparer.OrdinalIgnoreCase)
                .ThenBy(item => item.Family ?? string.Empty, StringComparer.OrdinalIgnoreCase)
                .ThenBy(item => item.TypeName ?? string.Empty, StringComparer.OrdinalIgnoreCase)
                .ToList();

            var result = new ReviewResult
            {
                File = safeFileLabel,
                CriteriaCount = criteriaSet.Count,
                TotalElements = groups.Sum(item => item.ElementCount),
                TotalGroups = groups.Count
            };

            int totalGroups = Math.Max(groups.Count, 1);
            for (int index = 0; index < groups.Count; index++)
            {
#if false
                progress?.Invoke(75d + (((double)index / totalGroups) * 25d), $"Family 적합성 판정 중 ({index + 1}/{groups.Count})");

 #endif
                progress?.Invoke(75d + (((double)index / totalGroups) * 25d), $"Family \uC801\uD569\uC131 \uD310\uC815 \uC911 ({index + 1}/{groups.Count})");
                UsageGroup group = groups[index];
                FilterRule matchedFilter = filterRules.FirstOrDefault(rule => IsFilterMatch(rule, group));

                string review;
                string reviewSource;
                string status;
                if (matchedFilter != null)
                {
                    review = matchedFilter.ReviewText;
                    reviewSource = "FILTER";
                    status = "WARN";
                    result.FilterCount += 1;
                }
                else if (criteriaSet.Contains(group.Key))
                {
                    review = matchReviewText;
                    reviewSource = "MATCH";
                    status = "OK";
                    result.MatchCount += 1;
                }
                else
                {
                    review = mismatchReviewText;
                    reviewSource = "MISMATCH";
                    status = "WARN";
                    result.MismatchCount += 1;
                }

                result.Rows.Add(new ReviewRow
                {
                    File = safeFileLabel,
                    Category = group.Category,
                    Family = group.Family,
                    TypeName = group.TypeName,
                    ElementCount = group.ElementCount,
                    Review = review,
                    ReviewSource = reviewSource,
                    Status = status
                });
            }

#if false

            progress?.Invoke(100d, "Family 적합성 검토 완료");
 #endif
            progress?.Invoke(100d, "Family \uC801\uD569\uC131 \uAC80\uD1A0 \uC644\uB8CC");
            result.FileSummaries.Add(new FileSummary
            {
                File = safeFileLabel,
                Total = result.TotalGroups,
                Issues = result.MismatchCount,
                Near = result.FilterCount,
                MatchCount = result.MatchCount,
                MismatchCount = result.MismatchCount,
                FilterCount = result.FilterCount,
                ElementCount = result.TotalElements,
                Status = "success",
#if false
                Reason = result.TotalGroups == 0 ? "집계 가능한 객체가 없습니다." : string.Empty
 #endif
                Reason = result.TotalGroups == 0 ? "\uC9D1\uACC4 \uAC00\uB2A5\uD55C \uAC1D\uCCB4\uAC00 \uC5C6\uC2B5\uB2C8\uB2E4." : string.Empty
            });
            return result;
        }

#if false

        public static DataTable BuildExportTable(IEnumerable<ReviewRow> rows, string emptyMessage = "집계 가능한 객체가 없습니다.")
        {
            var table = new DataTable("FamilySuitabilityReview");
            table.Columns.Add("Category");
            table.Columns.Add("Family");
            table.Columns.Add("Type");
            table.Columns.Add("No. of Elements", typeof(int));
            table.Columns.Add("Review");
            DataColumn reviewEnColumn = table.Columns.Add("__ReviewEn");
            reviewEnColumn.ExtendedProperties["ExcelHidden"] = true;
            DataColumn reviewKoColumn = table.Columns.Add("__ReviewKo");
            reviewKoColumn.ExtendedProperties["ExcelHidden"] = true;
            DataColumn statusColumn = table.Columns.Add("Status");
            statusColumn.ExtendedProperties["ExcelHidden"] = true;

            List<ReviewRow> source = (rows ?? Enumerable.Empty<ReviewRow>())
                .Where(row => row != null)
                .OrderBy(row => row.Category ?? string.Empty, StringComparer.OrdinalIgnoreCase)
                .ThenBy(row => row.Family ?? string.Empty, StringComparer.OrdinalIgnoreCase)
                .ThenBy(row => row.TypeName ?? string.Empty, StringComparer.OrdinalIgnoreCase)
                .ToList();

            if (source.Count == 0)
            {
                DataRow empty = table.NewRow();
                empty["Category"] = string.IsNullOrWhiteSpace(emptyMessage) ? "집계 가능한 객체가 없습니다." : emptyMessage.Trim();
                empty["__ReviewEn"] = LocalizeConfiguredReviewText(Convert.ToString(empty["Category"]), "en");
                empty["__ReviewKo"] = LocalizeConfiguredReviewText(Convert.ToString(empty["Category"]), "ko");
                empty["Status"] = "OK";
                table.Rows.Add(empty);
                return table;
            }

            foreach (ReviewRow row in source)
            {
                DataRow dataRow = table.NewRow();
                dataRow["Category"] = row.Category ?? string.Empty;
                dataRow["Family"] = row.Family ?? string.Empty;
                dataRow["Type"] = row.TypeName ?? string.Empty;
                dataRow["No. of Elements"] = row.ElementCount;
                dataRow["Review"] = row.Review ?? string.Empty;
                dataRow["__ReviewEn"] = LocalizeConfiguredReviewText(row.Review, "en");
                dataRow["__ReviewKo"] = LocalizeConfiguredReviewText(row.Review, "ko");
                dataRow["Status"] = row.Status ?? string.Empty;
                table.Rows.Add(dataRow);
            }

            return table;
        }

 #endif

        public static DataTable BuildExportTable(IEnumerable<ReviewRow> rows, string emptyMessage = "\uC9D1\uACC4 \uAC00\uB2A5\uD55C \uAC1D\uCCB4\uAC00 \uC5C6\uC2B5\uB2C8\uB2E4.")
        {
            var table = new DataTable("FamilySuitabilityReview");
            table.Columns.Add("Category");
            table.Columns.Add("Family");
            table.Columns.Add("Type");
            table.Columns.Add("No. of Elements", typeof(int));
            table.Columns.Add("Review");

            DataColumn reviewEnColumn = table.Columns.Add("__ReviewEn");
            reviewEnColumn.ExtendedProperties["ExcelHidden"] = true;

            DataColumn reviewKoColumn = table.Columns.Add("__ReviewKo");
            reviewKoColumn.ExtendedProperties["ExcelHidden"] = true;

            List<ReviewRow> source = (rows ?? Enumerable.Empty<ReviewRow>())
                .Where(row => row != null)
                .OrderBy(row => row.Category ?? string.Empty, StringComparer.OrdinalIgnoreCase)
                .ThenBy(row => row.Family ?? string.Empty, StringComparer.OrdinalIgnoreCase)
                .ThenBy(row => row.TypeName ?? string.Empty, StringComparer.OrdinalIgnoreCase)
                .ToList();

            if (source.Count == 0)
            {
                DataRow empty = table.NewRow();
                empty["Category"] = string.IsNullOrWhiteSpace(emptyMessage)
                    ? "\uC9D1\uACC4 \uAC00\uB2A5\uD55C \uAC1D\uCCB4\uAC00 \uC5C6\uC2B5\uB2C8\uB2E4."
                    : emptyMessage.Trim();
                empty["__ReviewEn"] = string.Empty;
                empty["__ReviewKo"] = string.Empty;
                table.Rows.Add(empty);
                return table;
            }

            foreach (ReviewRow row in source)
            {
                DataRow dataRow = table.NewRow();
                dataRow["Category"] = row.Category ?? string.Empty;
                dataRow["Family"] = row.Family ?? string.Empty;
                dataRow["Type"] = row.TypeName ?? string.Empty;
                dataRow["No. of Elements"] = row.ElementCount;
                dataRow["Review"] = row.Review ?? string.Empty;
                dataRow["__ReviewEn"] = LocalizeConfiguredReviewText(row.Review, "en");
                dataRow["__ReviewKo"] = LocalizeConfiguredReviewText(row.Review, "ko");
                table.Rows.Add(dataRow);
            }

            return table;
        }

        #if false
        private static string LocalizeConfiguredReviewText(string text, string locale)
        {
            string value = string.IsNullOrWhiteSpace(text) ? string.Empty : text.Trim();
            if (string.IsNullOrWhiteSpace(value)) return value;

            if (string.Equals(locale, "en", StringComparison.OrdinalIgnoreCase))
            {
                return ApplyConfiguredReviewReplacements(value, new[]
                {
                    new KeyValuePair<string, string>("검토 필요", "Needs review"),
                    new KeyValuePair<string, string>("기준 미일치", "Does not match criteria"),
                    new KeyValuePair<string, string>("기준 일치", "Matches criteria"),
                    new KeyValuePair<string, string>("불일치", "Mismatch"),
                    new KeyValuePair<string, string>("일치", "Match"),
                    new KeyValuePair<string, string>("기준 리스트", "criteria list"),
                    new KeyValuePair<string, string>("기준", "criteria"),
                    new KeyValuePair<string, string>("패밀리", "family"),
                    new KeyValuePair<string, string>("타입", "type"),
                    new KeyValuePair<string, string>("카테고리", "category"),
                    new KeyValuePair<string, string>("필터", "filter"),
                    new KeyValuePair<string, string>("누락", "missing"),
                    new KeyValuePair<string, string>("오류", "error")
                });
            }

            return ApplyConfiguredReviewReplacements(value, new[]
            {
                new KeyValuePair<string, string>("Does not match criteria", "기준 미일치"),
                new KeyValuePair<string, string>("Matches criteria", "기준 일치"),
                new KeyValuePair<string, string>("Needs review", "검토 필요"),
                new KeyValuePair<string, string>("Review required", "검토 필요"),
                new KeyValuePair<string, string>("criteria list", "기준 리스트"),
                new KeyValuePair<string, string>("criteria", "기준"),
                new KeyValuePair<string, string>("mismatch", "미일치"),
                new KeyValuePair<string, string>("match", "일치"),
                new KeyValuePair<string, string>("family", "패밀리"),
                new KeyValuePair<string, string>("type", "타입"),
                new KeyValuePair<string, string>("category", "카테고리"),
                new KeyValuePair<string, string>("filter", "필터"),
                new KeyValuePair<string, string>("missing", "누락"),
                new KeyValuePair<string, string>("error", "오류")
            });
        }

        #endif

        private static string LocalizeConfiguredReviewText(string text, string locale)
        {
            string value = string.IsNullOrWhiteSpace(text) ? string.Empty : text.Trim();
            if (string.IsNullOrWhiteSpace(value)) return value;

            if (string.Equals(locale, "en", StringComparison.OrdinalIgnoreCase))
            {
                return ApplyConfiguredReviewReplacements(value, new[]
                {
                    new KeyValuePair<string, string>("\uAC80\uD1A0 \uD544\uC694", "Needs review"),
                    new KeyValuePair<string, string>("\uAE30\uC900 \uBBF8\uC77C\uCE58", "Does not match criteria"),
                    new KeyValuePair<string, string>("\uAE30\uC900 \uC77C\uCE58", "Matches criteria"),
                    new KeyValuePair<string, string>("\uBD88\uC77C\uCE58", "Mismatch"),
                    new KeyValuePair<string, string>("\uC77C\uCE58", "Match"),
                    new KeyValuePair<string, string>("\uAE30\uC900 \uB9AC\uC2A4\uD2B8", "criteria list"),
                    new KeyValuePair<string, string>("\uAE30\uC900", "criteria"),
                    new KeyValuePair<string, string>("\uD328\uBC00\uB9AC", "family"),
                    new KeyValuePair<string, string>("\uD0C0\uC785", "type"),
                    new KeyValuePair<string, string>("\uCE74\uD14C\uACE0\uB9AC", "category"),
                    new KeyValuePair<string, string>("\uD544\uD130", "filter"),
                    new KeyValuePair<string, string>("\uB204\uB77D", "missing"),
                    new KeyValuePair<string, string>("\uC624\uB958", "error")
                });
            }

            return ApplyConfiguredReviewReplacements(value, new[]
            {
                new KeyValuePair<string, string>("Does not match criteria", "\uAE30\uC900 \uBBF8\uC77C\uCE58"),
                new KeyValuePair<string, string>("Matches criteria", "\uAE30\uC900 \uC77C\uCE58"),
                new KeyValuePair<string, string>("Needs review", "\uAC80\uD1A0 \uD544\uC694"),
                new KeyValuePair<string, string>("Review required", "\uAC80\uD1A0 \uD544\uC694"),
                new KeyValuePair<string, string>("criteria list", "\uAE30\uC900 \uB9AC\uC2A4\uD2B8"),
                new KeyValuePair<string, string>("criteria", "\uAE30\uC900"),
                new KeyValuePair<string, string>("mismatch", "\uBBF8\uC77C\uCE58"),
                new KeyValuePair<string, string>("match", "\uC77C\uCE58"),
                new KeyValuePair<string, string>("family", "\uD328\uBC00\uB9AC"),
                new KeyValuePair<string, string>("type", "\uD0C0\uC785"),
                new KeyValuePair<string, string>("category", "\uCE74\uD14C\uACE0\uB9AC"),
                new KeyValuePair<string, string>("filter", "\uD544\uD130"),
                new KeyValuePair<string, string>("missing", "\uB204\uB77D"),
                new KeyValuePair<string, string>("error", "\uC624\uB958")
            });
        }

        private static string ApplyConfiguredReviewReplacements(string text, IEnumerable<KeyValuePair<string, string>> replacements)
        {
            string result = text ?? string.Empty;
            if (string.IsNullOrWhiteSpace(result) || replacements == null) return result;

            foreach (KeyValuePair<string, string> replacement in replacements)
            {
                if (string.IsNullOrWhiteSpace(replacement.Key)) continue;

                result = Regex.Replace(
                    result,
                    Regex.Escape(replacement.Key),
                    replacement.Value ?? string.Empty,
                    RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
            }

            return result;
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
            if (element.Category == null) return false;
            if (string.IsNullOrWhiteSpace(element.Category.Name)) return false;
            string categoryName = element.Category.Name.Trim();
            if (categoryName.IndexOf("Runs", StringComparison.OrdinalIgnoreCase) >= 0) return false;
            if (element.Category.CategoryType == CategoryType.Internal) return false;
            if (element.Category.CategoryType == CategoryType.Annotation) return false;
            if (IsExcludedExternalReferenceElement(element)) return false;
            if (element.GetTypeId() == ElementId.InvalidElementId) return false;

            if (element is View) return false;
            if (element is ViewSheet) return false;
            if (element is Viewport) return false;
            if (element is Level) return false;
            if (element is Grid) return false;
            if (element is ReferencePlane) return false;
            if (element is CurveElement) return false;
            if (element is SketchPlane) return false;
            if (element is Autodesk.Revit.DB.Group) return false;
            if (element is AssemblyInstance) return false;
            if (element is RevitLinkInstance) return false;
            if (element is ImportInstance) return false;
            if (element is BasePoint) return false;
            if (element is ProjectInfo) return false;
            if (element is Room) return false;
            if (element is Area) return false;
            if (element is Space) return false;
            if (element is MEPSystem) return false;

            int categoryId = element.Category.Id.IntegerValue;
            if (categoryId == (int)BuiltInCategory.OST_Levels) return false;
            if (categoryId == (int)BuiltInCategory.OST_Grids) return false;
            if (categoryId == (int)BuiltInCategory.OST_DetailComponents) return false;
            if (categoryId == (int)BuiltInCategory.OST_PointClouds) return false;
            if (categoryId == (int)BuiltInCategory.OST_RvtLinks) return false;
            if (categoryId == (int)BuiltInCategory.OST_Cameras) return false;
            if (categoryId == (int)BuiltInCategory.OST_VolumeOfInterest) return false;
            if (categoryId == (int)BuiltInCategory.OST_SectionBox) return false;
            if (categoryId == (int)BuiltInCategory.OST_Rooms) return false;
            if (categoryId == (int)BuiltInCategory.OST_Areas) return false;
            if (categoryId == (int)BuiltInCategory.OST_MEPSpaces) return false;
            if (categoryId == (int)BuiltInCategory.OST_ConduitRun) return false;
            return true;
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

        private static UsageGroup TryBuildUsageGroup(Document doc, Element element)
        {
            try
            {
                ElementId typeId = element.GetTypeId();
                if (typeId == null || typeId == ElementId.InvalidElementId) return null;

                ElementType elementType = doc.GetElement(typeId) as ElementType;
                if (elementType == null) return null;

                string category = CleanDisplayText(element.Category != null ? element.Category.Name : string.Empty);
                string family = ResolveFamilyName(element, elementType);
                string typeName = ResolveTypeName(element, elementType);
                if (string.IsNullOrWhiteSpace(category) || string.IsNullOrWhiteSpace(family) || string.IsNullOrWhiteSpace(typeName))
                {
                    return null;
                }

                return new UsageGroup
                {
                    Category = category,
                    Family = family,
                    TypeName = typeName,
                    NormalizedFamily = NormalizeToken(family),
                    NormalizedTypeName = NormalizeToken(typeName),
                    Key = BuildKey(category, family, typeName)
                };
            }
            catch
            {
                return null;
            }
        }

        private static string ResolveFamilyName(Element element, ElementType elementType)
        {
            FamilyInstance familyInstance = element as FamilyInstance;
            if (familyInstance != null)
            {
                FamilySymbol symbol = familyInstance.Symbol;
                if (symbol != null)
                {
                    string familyName = CleanDisplayText(symbol.FamilyName);
                    if (!string.IsNullOrWhiteSpace(familyName)) return familyName;
                    if (symbol.Family != null)
                    {
                        familyName = CleanDisplayText(symbol.Family.Name);
                        if (!string.IsNullOrWhiteSpace(familyName)) return familyName;
                    }
                }
            }

            string systemFamilyName = CleanDisplayText(elementType.FamilyName);
            if (!string.IsNullOrWhiteSpace(systemFamilyName)) return systemFamilyName;

            return CleanDisplayText(elementType.Name);
        }

        private static string ResolveTypeName(Element element, ElementType elementType)
        {
            string typeName = CleanDisplayText(elementType.Name);
            if (!string.IsNullOrWhiteSpace(typeName)) return typeName;
            return CleanDisplayText(element.Name);
        }

        private static HashSet<string> BuildCriteriaSet(IEnumerable<CriteriaRule> rules)
        {
            var set = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (CriteriaRule rule in rules ?? Enumerable.Empty<CriteriaRule>())
            {
                if (rule == null) continue;
                string key = BuildKey(rule.Category, rule.Family, rule.TypeName);
                if (!string.IsNullOrWhiteSpace(key))
                {
                    set.Add(key);
                }
            }
            return set;
        }

        private static List<FilterRule> NormalizeFilterRules(IEnumerable<FilterRule> rules)
        {
            return (rules ?? Enumerable.Empty<FilterRule>())
                .Where(rule => rule != null)
                .Select(rule => new FilterRule
                {
                    Target = NormalizeFilterTarget(rule.Target),
                    Keyword = CleanDisplayText(rule.Keyword),
                    ReviewText = CleanDisplayText(rule.ReviewText)
                })
                .Where(rule => !string.IsNullOrWhiteSpace(rule.Keyword) && !string.IsNullOrWhiteSpace(rule.ReviewText))
                .ToList();
        }

        private static bool IsFilterMatch(FilterRule rule, UsageGroup group)
        {
            if (rule == null || group == null) return false;
            string keyword = NormalizeToken(rule.Keyword);
            if (string.IsNullOrWhiteSpace(keyword)) return false;

            string target = NormalizeFilterTarget(rule.Target);
            if (string.Equals(target, "family", StringComparison.OrdinalIgnoreCase))
            {
                return group.NormalizedFamily.Contains(keyword);
            }

            if (string.Equals(target, "type", StringComparison.OrdinalIgnoreCase))
            {
                return group.NormalizedTypeName.Contains(keyword);
            }

            return group.NormalizedFamily.Contains(keyword) || group.NormalizedTypeName.Contains(keyword);
        }

        private static string NormalizeFilterTarget(string target)
        {
            string normalized = NormalizeToken(target);
            if (normalized == "family") return "family";
            if (normalized == "type") return "type";
            return "familyOrType";
        }

        private static string BuildKey(string category, string family, string typeName)
        {
            string normalizedCategory = NormalizeToken(category);
            string normalizedFamily = NormalizeToken(family);
            string normalizedType = NormalizeToken(typeName);
            if (string.IsNullOrWhiteSpace(normalizedCategory) || string.IsNullOrWhiteSpace(normalizedFamily) || string.IsNullOrWhiteSpace(normalizedType))
            {
                return string.Empty;
            }
            return normalizedCategory + "|" + normalizedFamily + "|" + normalizedType;
        }

        private static string CleanDisplayText(string value)
        {
            if (string.IsNullOrWhiteSpace(value)) return string.Empty;
            string normalized = value.Normalize(NormalizationForm.FormKC).Trim();
            return MultiWhitespace.Replace(normalized, " ").Trim();
        }

        private static string NormalizeToken(string value)
        {
            return CleanDisplayText(value).ToLowerInvariant();
        }
    }
}
