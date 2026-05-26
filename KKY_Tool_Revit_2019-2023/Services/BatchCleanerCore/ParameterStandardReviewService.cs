using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;
using Autodesk.Revit.DB;

namespace KKY_Tool_Revit.Services
{
    public static class ParameterStandardReviewService
    {
        public const string ItemLabel = "\uC18D\uC131 \uBAA8\uC218 \uAC80\uD1A0 \uC624\uB958";
        private static readonly Regex MultiWhitespace = new Regex(@"\s+", RegexOptions.Compiled);

        public sealed class Settings
        {
            public List<CriteriaRule> CriteriaRules { get; set; } = new List<CriteriaRule>();
            public bool HasAllowedElementScope { get; set; }
            public List<int> AllowedElementIds { get; set; } = new List<int>();
            public string CommonTargetFilterText { get; set; } = string.Empty;
            public string CommonExcludeTargetFilterText { get; set; } = string.Empty;
            public List<string> ExtraParameterNames { get; set; } = new List<string>();
        }

        public sealed class CriteriaRule
        {
            public string ParameterName { get; set; } = string.Empty;
            public string SheetName { get; set; } = string.Empty;
            public string HeaderParameterName { get; set; } = string.Empty;
            public List<string> AllowedValues { get; set; } = new List<string>();
            public bool AllowBlank { get; set; }
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
            public string Parameter { get; set; } = string.Empty;
            public string CurrentValue { get; set; } = string.Empty;
            public string AllowedValues { get; set; } = string.Empty;
            public string Solutions { get; set; } = string.Empty;
            public Dictionary<string, string> ExtraParams { get; set; } = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        }

        public sealed class FileSummary
        {
            public string File { get; set; } = string.Empty;
            public int TargetElementCount { get; set; }
            public int ParameterCount { get; set; }
            public int TotalReviewed { get; set; }
            public int ErrorCount { get; set; }
            public int OkCount { get; set; }
            public int BlankAllowedCount { get; set; }
            public string Status { get; set; } = "pending";
            public string Reason { get; set; } = string.Empty;
        }

        public sealed class ReviewResult
        {
            public string File { get; set; } = string.Empty;
            public int TargetElementCount { get; set; }
            public int ParameterCount { get; set; }
            public int TotalReviewed { get; set; }
            public int ErrorCount { get; set; }
            public int OkCount { get; set; }
            public List<ReviewRow> Rows { get; set; } = new List<ReviewRow>();
            public List<FileSummary> FileSummaries { get; set; } = new List<FileSummary>();
        }

        private sealed class PreparedRule
        {
            public string ParameterName { get; set; } = string.Empty;
            public List<string> AllowedValues { get; set; } = new List<string>();
            public HashSet<string> AllowedSet { get; set; } = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            public bool AllowBlank { get; set; }
            public string AllowedValuesText { get; set; } = string.Empty;
        }

        private sealed class TypeInfo
        {
            public string Category { get; set; } = string.Empty;
            public string Family { get; set; } = string.Empty;
            public string TypeName { get; set; } = string.Empty;
        }

        private sealed class ReviewValueCache
        {
            private readonly Document _doc;
            private readonly Dictionary<string, ModelParameterExtractionService.ElementParameterValueInfo> _parameterValues =
                new Dictionary<string, ModelParameterExtractionService.ElementParameterValueInfo>(StringComparer.OrdinalIgnoreCase);
            private readonly Dictionary<int, TypeInfo> _typeInfos =
                new Dictionary<int, TypeInfo>();

            public ReviewValueCache(Document doc)
            {
                _doc = doc;
            }

            public TypeInfo GetTypeInfo(Element element)
            {
                if (_doc == null || element == null)
                {
                    return new TypeInfo();
                }

                int elementId = element.Id?.CompatIntegerValue() ?? 0;
                TypeInfo info;
                if (!_typeInfos.TryGetValue(elementId, out info))
                {
                    info = new TypeInfo
                    {
                        Category = ModelParameterExtractionService.GetElementCategoryName(element),
                        Family = ModelParameterExtractionService.GetElementFamilyName(_doc, element),
                        TypeName = ModelParameterExtractionService.GetElementTypeName(_doc, element)
                    };
                    _typeInfos[elementId] = info;
                }

                return info;
            }

            public ModelParameterExtractionService.ElementParameterValueInfo GetValueInfo(Element element, string parameterName)
            {
                if (_doc == null || element == null || string.IsNullOrWhiteSpace(parameterName))
                {
                    return new ModelParameterExtractionService.ElementParameterValueInfo();
                }

                string normalizedName = CleanText(parameterName);
                TypeInfo typeInfo = GetTypeInfo(element);
                if (IsPseudoParameter(normalizedName, "Category"))
                {
                    return CreatePseudoValue(typeInfo.Category);
                }
                if (IsPseudoParameter(normalizedName, "Family") || IsPseudoParameter(normalizedName, "FamilyName"))
                {
                    return CreatePseudoValue(typeInfo.Family);
                }
                if (IsPseudoParameter(normalizedName, "Type") || IsPseudoParameter(normalizedName, "TypeName"))
                {
                    return CreatePseudoValue(typeInfo.TypeName);
                }

                int elementId = element.Id?.CompatIntegerValue() ?? 0;
                string key = elementId.ToString(CultureInfo.InvariantCulture) + "\u001f" + normalizedName;
                ModelParameterExtractionService.ElementParameterValueInfo info;
                if (!_parameterValues.TryGetValue(key, out info))
                {
                    info = ModelParameterExtractionService.GetElementParameterValueInfo(_doc, element, normalizedName)
                           ?? new ModelParameterExtractionService.ElementParameterValueInfo();
                    _parameterValues[key] = info;
                }

                return info;
            }

            private static bool IsPseudoParameter(string actual, string expected)
            {
                return string.Equals(actual, expected, StringComparison.OrdinalIgnoreCase);
            }

            private static ModelParameterExtractionService.ElementParameterValueInfo CreatePseudoValue(string value)
            {
                return new ModelParameterExtractionService.ElementParameterValueInfo
                {
                    HasParameter = true,
                    ValueText = value ?? string.Empty,
                    DataTypeToken = "Text"
                };
            }
        }

        public static ReviewResult RunOnDocument(Document doc, string fileLabel, Settings settings, Action<double, string> progress = null)
        {
            if (doc == null) throw new ArgumentNullException(nameof(doc));
            if (settings == null) throw new ArgumentNullException(nameof(settings));

            string safeFileLabel = string.IsNullOrWhiteSpace(fileLabel) ? (doc.Title ?? string.Empty) : fileLabel.Trim();
            List<PreparedRule> rules = PrepareRules(settings.CriteriaRules);
            if (rules.Count == 0)
            {
                throw new InvalidOperationException("\uAE30\uC900 \uC5D1\uC140\uC5D0\uC11C \uC0AC\uC6A9\uD560 \uD30C\uB77C\uBBF8\uD130 \uAE30\uC900\uAC12\uC744 \uCC3E\uC9C0 \uBABB\uD588\uC2B5\uB2C8\uB2E4.");
            }

            HashSet<int> allowedElementIds = new HashSet<int>((settings.AllowedElementIds ?? Enumerable.Empty<int>()).Where(id => id > 0));
            bool hasCommonScopeFilter = settings.HasAllowedElementScope;
            List<string> extraParameterNames = NormalizeExtraParameterNames(settings.ExtraParameterNames);
            ReviewValueCache valueCache = new ReviewValueCache(doc);

            progress?.Invoke(5d, "\uAC80\uD1A0 \uB300\uC0C1 \uAC1D\uCCB4 \uC218\uC9D1 \uC911");
            List<Element> candidates = ModelParameterExtractionService.GetExtractableElements(doc)
                .Where(element => element != null)
                .ToList();

            List<Element> targetElements = candidates
                .Where(element => !hasCommonScopeFilter || (element.Id != null && allowedElementIds.Contains(element.Id.CompatIntegerValue())))
                .ToList();

            var result = new ReviewResult
            {
                File = safeFileLabel,
                TargetElementCount = targetElements.Count,
                ParameterCount = rules.Count,
                TotalReviewed = 0
            };

            int totalElements = Math.Max(targetElements.Count, 1);
            for (int index = 0; index < targetElements.Count; index++)
            {
                Element element = targetElements[index];
                TypeInfo typeInfo = null;

                foreach (PreparedRule rule in rules)
                {
                    ModelParameterExtractionService.ElementParameterValueInfo valueInfo = valueCache.GetValueInfo(element, rule.ParameterName);
                    if (!valueInfo.HasParameter)
                    {
                        continue;
                    }

                    result.TotalReviewed++;
                    string currentValue = CleanText(valueInfo.ValueText);
                    bool isBlank = string.IsNullOrWhiteSpace(currentValue);
                    bool passes = (isBlank && rule.AllowBlank) || (!isBlank && rule.AllowedSet.Contains(currentValue));

                    if (passes)
                    {
                        result.OkCount++;
                        continue;
                    }

                    result.ErrorCount++;
                    typeInfo = typeInfo ?? valueCache.GetTypeInfo(element);
                    result.Rows.Add(new ReviewRow
                    {
                        File = safeFileLabel,
                        Item = ItemLabel,
                        Id = (element.Id?.CompatIntegerValue() ?? 0).ToString(CultureInfo.InvariantCulture),
                        Name = typeInfo.TypeName,
                        Result = "Error",
                        Content = BuildContentMessage(rule, valueInfo, currentValue),
                        Etc = string.Empty,
                        Category = typeInfo.Category,
                        Family = typeInfo.Family,
                        Parameter = rule.ParameterName,
                        CurrentValue = currentValue,
                        AllowedValues = rule.AllowedValuesText,
                        Solutions = "\uAE30\uC900 \uC5D1\uC140\uC758 \uD5C8\uC6A9\uAC12\uC5D0 \uB9DE\uAC8C \uD30C\uB77C\uBBF8\uD130 \uAC12\uC744 \uC218\uC815\uD558\uAC70\uB098, \uAE30\uC900 \uBAA9\uB85D\uC744 \uAC31\uC2E0\uD558\uC138\uC694.",
                        ExtraParams = ReadExtraParamValues(element, extraParameterNames, valueCache)
                    });
                }

                if (index == 0 || index == targetElements.Count - 1 || index % 120 == 0)
                {
                    progress?.Invoke(5d + (((double)(index + 1) / totalElements) * 95d), $"\uC18D\uC131 \uAE30\uC900\uAC12 \uAC80\uD1A0 \uC911 ({index + 1}/{targetElements.Count})");
                }
            }

            if (result.ErrorCount == 0)
            {
                result.Rows.Add(new ReviewRow
                {
                    File = safeFileLabel,
                    Item = BuildItemText(0),
                    Result = "OK",
                    Content = BuildOkMessage(result),
                    ExtraParams = BuildEmptyExtraParamValues(extraParameterNames)
                });
            }
            else
            {
                string itemText = BuildItemText(result.ErrorCount);
                foreach (ReviewRow row in result.Rows.Where(row => row != null))
                {
                    row.Item = itemText;
                }
            }

            result.FileSummaries.Add(new FileSummary
            {
                File = safeFileLabel,
                TargetElementCount = result.TargetElementCount,
                ParameterCount = result.ParameterCount,
                TotalReviewed = result.TotalReviewed,
                ErrorCount = result.ErrorCount,
                OkCount = result.OkCount,
                BlankAllowedCount = rules.Count(rule => rule.AllowBlank),
                Status = "success",
                Reason = BuildSummaryReason(result, hasCommonScopeFilter)
            });

            progress?.Invoke(100d, "\uC18D\uC131 \uBAA8\uC218 \uAC80\uD1A0 \uC644\uB8CC");
            return result;
        }

        public static DataTable BuildExportTable(IEnumerable<ReviewRow> rows)
        {
            var table = new DataTable("ParameterStandardReview");
            table.Columns.Add("\uD56D\uBAA9");
            table.Columns.Add("ID");
            table.Columns.Add("NAME");
            table.Columns.Add("\uC624\uB958");
            table.Columns.Add("\uB0B4\uC6A9");
            table.Columns.Add("\uBE44\uACE0");
            table.Columns.Add("Category");
            table.Columns.Add("Family");

            List<ReviewRow> source = (rows ?? Enumerable.Empty<ReviewRow>())
                .Where(row => row != null)
                .Where(IsErrorRow)
                .ToList();

            string itemText = BuildExportItemText(source.Count);

            foreach (ReviewRow row in source)
            {
                DataRow dataRow = table.NewRow();
                dataRow["\uD56D\uBAA9"] = itemText;
                dataRow["ID"] = FormatExcelId(row.Id);
                dataRow["NAME"] = row.Name ?? string.Empty;
                dataRow["\uC624\uB958"] = "\uC624\uB958";
                dataRow["\uB0B4\uC6A9"] = row.Content ?? string.Empty;
                dataRow["\uBE44\uACE0"] = string.Empty;
                dataRow["Category"] = row.Category ?? string.Empty;
                dataRow["Family"] = row.Family ?? string.Empty;
                table.Rows.Add(dataRow);
            }

            return table;
        }

        private static List<PreparedRule> PrepareRules(IEnumerable<CriteriaRule> rules)
        {
            var result = new List<PreparedRule>();
            var seenParameters = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            foreach (CriteriaRule rule in rules ?? Enumerable.Empty<CriteriaRule>())
            {
                string parameterName = CleanText(rule?.ParameterName);
                if (string.IsNullOrWhiteSpace(parameterName)) continue;
                if (!seenParameters.Add(parameterName)) continue;

                var allowedValues = new List<string>();
                var allowedSet = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                foreach (string raw in rule.AllowedValues ?? Enumerable.Empty<string>())
                {
                    string value = CleanText(raw);
                    if (string.IsNullOrWhiteSpace(value)) continue;
                    if (allowedSet.Add(value)) allowedValues.Add(value);
                }

                bool allowBlank = rule.AllowBlank;
                if (!allowBlank && allowedValues.Count == 0) continue;

                result.Add(new PreparedRule
                {
                    ParameterName = parameterName,
                    AllowedValues = allowedValues,
                    AllowedSet = allowedSet,
                    AllowBlank = allowBlank,
                    AllowedValuesText = BuildAllowedValueText(allowedValues, allowBlank)
                });
            }

            return result;
        }

        private static string BuildContentMessage(PreparedRule rule, ModelParameterExtractionService.ElementParameterValueInfo valueInfo, string currentValue)
        {
            string parameterName = rule?.ParameterName ?? string.Empty;
            string valueText = currentValue ?? string.Empty;
            return $"[{parameterName}] \uD30C\uB77C\uBBF8\uD130 \uAE30\uC900(\uC624/\uD0C8\uC790) \uC624\uB958: {valueText}";
        }

        private static string BuildItemText(int errorCount)
        {
            return errorCount <= 0
                ? "\uC18D\uC131 \uBAA8\uC218 \uAC80\uD1A0 \uC815\uC0C1"
                : $"{ItemLabel} {errorCount.ToString(CultureInfo.InvariantCulture)}\uAC74";
        }

        private static string BuildExportItemText(int errorCount)
        {
            return $"Parameter \uBCC4 \uC18D\uC131\uC815\uBCF4 \uAE30\uC900 \uC624\uB958 Check( {Math.Max(errorCount, 0).ToString(CultureInfo.InvariantCulture)}\uAC74)";
        }

        private static bool IsErrorRow(ReviewRow row)
        {
            return row != null && string.Equals(row.Result, "Error", StringComparison.OrdinalIgnoreCase);
        }

        private static string FormatExcelId(string id)
        {
            string value = (id ?? string.Empty).Trim();
            if (string.IsNullOrWhiteSpace(value)) return string.Empty;
            return value.EndsWith(",", StringComparison.Ordinal) ? value : value + ",";
        }

        private static string BuildOkMessage(ReviewResult result)
        {
            int targets = result?.TargetElementCount ?? 0;
            if (targets <= 0)
            {
                return "\uAC80\uD1A0 \uB300\uC0C1 \uAC1D\uCCB4\uAC00 \uC5C6\uC2B5\uB2C8\uB2E4.";
            }

            int reviewed = result?.TotalReviewed ?? 0;
            if (reviewed <= 0)
            {
                return "검토할 파라미터가 있는 대상 객체가 없습니다.";
            }

            return $"{reviewed.ToString(CultureInfo.InvariantCulture)}\uAC74\uC758 \uD30C\uB77C\uBBF8\uD130 \uAC12\uC774 \uAE30\uC900\uAC12\uACFC \uC77C\uCE58\uD569\uB2C8\uB2E4.";
        }

        private static string BuildSummaryReason(ReviewResult result, bool hasScopeFilter)
        {
            if (result == null)
            {
                return "\uAC80\uD1A0 \uACB0\uACFC\uAC00 \uC5C6\uC2B5\uB2C8\uB2E4.";
            }

            if (result.TargetElementCount <= 0)
            {
                return "\uAC80\uD1A0 \uB300\uC0C1 \uAC1D\uCCB4\uAC00 \uC5C6\uC2B5\uB2C8\uB2E4.";
            }

            if (result.ErrorCount <= 0)
            {
                return BuildOkMessage(result);
            }

            string scopeText = hasScopeFilter ? " / \uACF5\uD1B5 \uD544\uD130 \uC801\uC6A9" : string.Empty;
            return $"\uC18D\uC131 \uAE30\uC900\uAC12 \uBD88\uC77C\uCE58\uAC00 {result.ErrorCount.ToString(CultureInfo.InvariantCulture)}\uAC74 \uBC1C\uACAC\uB418\uC5C8\uC2B5\uB2C8\uB2E4.{scopeText}";
        }

        private static string BuildAllowedValueText(IList<string> allowedValues, bool allowBlank)
        {
            var displayValues = new List<string>();
            if (allowBlank) displayValues.Add("(\uACF5\uB780)");
            displayValues.AddRange((allowedValues ?? Enumerable.Empty<string>()).Where(value => !string.IsNullOrWhiteSpace(value)));

            const int maxValues = 24;
            if (displayValues.Count <= maxValues)
            {
                return string.Join(", ", displayValues);
            }

            return string.Join(", ", displayValues.Take(maxValues))
                   + $" \uC678 {displayValues.Count - maxValues}\uAC1C";
        }

        private static List<string> NormalizeExtraParameterNames(IEnumerable<string> extraParameterNames)
        {
            return (extraParameterNames ?? Enumerable.Empty<string>())
                .Select(CleanText)
                .Where(name => !string.IsNullOrWhiteSpace(name))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        private static Dictionary<string, string> ReadExtraParamValues(Element element, IEnumerable<string> extraParameterNames, ReviewValueCache valueCache)
        {
            var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            if (element == null || valueCache == null) return result;

            foreach (string name in extraParameterNames ?? Enumerable.Empty<string>())
            {
                if (string.IsNullOrWhiteSpace(name)) continue;
                result[name] = valueCache.GetValueInfo(element, name).ValueText ?? string.Empty;
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
                    string normalized = CleanText(key);
                    if (string.IsNullOrWhiteSpace(normalized)) continue;
                    if (seen.Add(normalized)) result.Add(normalized);
                }
            }

            return result;
        }

        private static string CleanText(string value)
        {
            string text = value ?? string.Empty;
            if (string.IsNullOrWhiteSpace(text)) return string.Empty;
            return MultiWhitespace.Replace(text.Trim(), " ");
        }
    }
}
