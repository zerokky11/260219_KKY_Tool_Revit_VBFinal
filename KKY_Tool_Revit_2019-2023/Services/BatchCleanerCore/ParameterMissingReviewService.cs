using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.Linq;
using System.Text;
using Autodesk.Revit.DB;
using KKY_Tool_Revit.Models;

namespace KKY_Tool_Revit.Services
{
    public static class ParameterMissingReviewService
    {
        public const string ItemLabel = "속성누락검토 오류";

        private enum ParameterReviewState
        {
            HasValue = 0,
            EmptyValue = 1,
            ParameterNotFound = 2
        }

        public sealed class Settings
        {
            public List<string> ParameterNames { get; set; } = new List<string>();
            public ElementParameterUpdateSettings TargetFilter { get; set; } = new ElementParameterUpdateSettings();
            public bool ExcludeTargetFilterMatches { get; set; }
            public bool HasAllowedElementScope { get; set; }
            public List<int> AllowedElementIds { get; set; } = new List<int>();
            public string CommonTargetFilterText { get; set; } = string.Empty;
            public string CommonExcludeTargetFilterText { get; set; } = string.Empty;
            public List<string> ExtraParameterNames { get; set; } = new List<string>();
            public List<MissingRule> ExceptionRules { get; set; } = new List<MissingRule>();
        }

        public sealed class MissingRule
        {
            public bool Enabled { get; set; } = true;
            public string ParameterName { get; set; } = string.Empty;
            public ParameterConditionCombination CombinationMode { get; set; } = ParameterConditionCombination.Or;
            public List<ElementParameterCondition> Conditions { get; set; } = new List<ElementParameterCondition>();

            public bool HasConfiguredConditions()
            {
                return !string.IsNullOrWhiteSpace(ParameterName)
                    && Conditions != null
                    && Conditions.Any(condition => condition != null && condition.IsConfigured() && (condition.Enabled || !string.IsNullOrWhiteSpace(condition.ParameterName)));
            }

            public MissingRule Clone()
            {
                return new MissingRule
                {
                    Enabled = Enabled,
                    ParameterName = ParameterName,
                    CombinationMode = CombinationMode,
                    Conditions = (Conditions ?? Enumerable.Empty<ElementParameterCondition>())
                        .Where(condition => condition != null)
                        .Select(condition => condition.Clone())
                        .ToList()
                };
            }

            public string BuildSummary()
            {
                List<string> conditionTexts = (Conditions ?? Enumerable.Empty<ElementParameterCondition>())
                    .Where(condition => condition != null && condition.IsConfigured() && (condition.Enabled || !string.IsNullOrWhiteSpace(condition.ParameterName)))
                    .Select(condition => condition.ToString())
                    .Where(text => !string.IsNullOrWhiteSpace(text))
                    .ToList();

                if (conditionTexts.Count == 0) return string.Empty;

                string joiner = CombinationMode == ParameterConditionCombination.Or ? " OR " : " AND ";
                return string.Join(joiner, conditionTexts);
            }
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
            public string Solutions { get; set; } = string.Empty;
            public string Comments
            {
                get { return Solutions; }
                set { Solutions = value ?? string.Empty; }
            }
            public Dictionary<string, string> ExtraParams { get; set; } = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        }

        public sealed class FileSummary
        {
            public string File { get; set; } = string.Empty;
            public int TargetElementCount { get; set; }
            public int ParameterCount { get; set; }
            public int TotalReviewed { get; set; }
            public int ErrorCount { get; set; }
            public int IgnoredCount { get; set; }
            public int OkCount { get; set; }
            public int TargetConditionCount { get; set; }
            public int ExceptionRuleCount { get; set; }
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
            public int IgnoredCount { get; set; }
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

        public static ReviewResult RunOnDocument(Document doc, string fileLabel, Settings settings, Action<double, string> progress = null)
        {
            if (doc == null) throw new ArgumentNullException(nameof(doc));
            if (settings == null) throw new ArgumentNullException(nameof(settings));

            string safeFileLabel = string.IsNullOrWhiteSpace(fileLabel) ? (doc.Title ?? string.Empty) : fileLabel.Trim();
            List<string> parameterNames = NormalizeParameterNames(settings.ParameterNames);
            if (parameterNames.Count == 0)
            {
                throw new InvalidOperationException("검토할 파라미터명을 1개 이상 선택하세요.");
            }

            ElementParameterUpdateSettings targetFilter = settings.TargetFilter?.Clone() ?? new ElementParameterUpdateSettings();
            List<ElementParameterCondition> targetConditions = GetEnabledConditions(targetFilter);
            string commonTargetFilterText = (settings.CommonTargetFilterText ?? string.Empty).Trim();
            string commonExcludeTargetFilterText = (settings.CommonExcludeTargetFilterText ?? string.Empty).Trim();
            bool hasCommonScopeFilter = settings.HasAllowedElementScope;
            HashSet<int> allowedElementIds = new HashSet<int>((settings.AllowedElementIds ?? Enumerable.Empty<int>()).Where(id => id > 0));
            List<MissingRule> exceptionRules = NormalizeExceptionRules(settings.ExceptionRules, parameterNames);
            List<string> extraParameterNames = NormalizeExtraParameterNames(settings.ExtraParameterNames);

            progress?.Invoke(5d, "검토 대상 수집 중");
            List<Element> candidates = ModelParameterExtractionService.GetExtractableElements(doc)
                .Where(element => element != null)
                .ToList();

            List<Element> targetElements = new List<Element>();
            int candidateTotal = Math.Max(candidates.Count, 1);
            for (int index = 0; index < candidates.Count; index++)
            {
                Element element = candidates[index];
                bool matchesLocalTarget = true;
                if (targetConditions.Count > 0)
                {
                    bool conditionMatched = MatchesConditions(doc, element, targetConditions, targetFilter.CombinationMode);
                    matchesLocalTarget = settings.ExcludeTargetFilterMatches ? !conditionMatched : conditionMatched;
                }
                bool matchesCommonScope = !hasCommonScopeFilter || (element?.Id != null && allowedElementIds.Contains(element.Id.IntegerValue));
                if (matchesLocalTarget && matchesCommonScope)
                {
                    targetElements.Add(element);
                }

                if (index == 0 || index == candidates.Count - 1 || index % 250 == 0)
                {
                    progress?.Invoke(5d + (((double)(index + 1) / candidateTotal) * 15d), $"검토 대상 수집 중 ({index + 1}/{candidates.Count})");
                }
            }

            var result = new ReviewResult
            {
                File = safeFileLabel,
                TargetElementCount = targetElements.Count,
                ParameterCount = parameterNames.Count,
                TotalReviewed = targetElements.Count * parameterNames.Count
            };

            int elementTotal = Math.Max(targetElements.Count, 1);
            for (int index = 0; index < targetElements.Count; index++)
            {
                Element element = targetElements[index];
                TypeInfo typeInfo = ResolveTypeInfo(doc, element);

                foreach (string parameterName in parameterNames)
                {
                    ParameterReviewState reviewState = GetParameterReviewState(doc, element, parameterName);
                    if (reviewState == ParameterReviewState.HasValue)
                    {
                        result.OkCount++;
                        continue;
                    }

                    MissingRule matchedRule;
                    if (ShouldIgnoreMissing(doc, element, parameterName, exceptionRules, out matchedRule))
                    {
                        result.IgnoredCount++;
                        continue;
                    }

                    result.ErrorCount++;
                    result.Rows.Add(new ReviewRow
                    {
                        File = safeFileLabel,
                        Item = ItemLabel,
                        Id = (element?.Id?.IntegerValue ?? 0).ToString(CultureInfo.InvariantCulture),
                        Name = typeInfo.TypeName ?? string.Empty,
                        Result = "Error",
                        Content = BuildContentMessage(parameterName, reviewState),
                        Etc = string.Empty,
                        Category = typeInfo.Category,
                        Family = typeInfo.Family,
                        Solutions = BuildSolutionsMessage(reviewState),
                        ExtraParams = ReadExtraParamValues(doc, element, extraParameterNames)
                    });
                }

                if (index == 0 || index == targetElements.Count - 1 || index % 120 == 0)
                {
                    progress?.Invoke(20d + (((double)(index + 1) / elementTotal) * 80d), $"누락 값 검토 중 ({index + 1}/{targetElements.Count})");
                }
            }

            if (result.ErrorCount == 0)
            {
                result.Rows.Add(new ReviewRow
                {
                    File = safeFileLabel,
                    Item = BuildItemText(0),
                    Id = string.Empty,
                    Name = string.Empty,
                    Result = "OK",
                    Content = BuildOkMessage(result),
                    Etc = string.Empty,
                    Category = string.Empty,
                    Family = string.Empty,
                    Solutions = string.Empty,
                    ExtraParams = BuildEmptyExtraParamValues(extraParameterNames)
                });
            }
            else
            {
                string errorItemText = BuildItemText(result.ErrorCount);
                foreach (ReviewRow row in result.Rows.Where(row => row != null))
                {
                    row.Item = errorItemText;
                }
            }

            progress?.Invoke(100d, "파라미터 누락 검토 완료");

            result.FileSummaries.Add(new FileSummary
            {
                File = safeFileLabel,
                TargetElementCount = result.TargetElementCount,
                ParameterCount = result.ParameterCount,
                TotalReviewed = result.TotalReviewed,
                ErrorCount = result.ErrorCount,
                IgnoredCount = result.IgnoredCount,
                OkCount = result.OkCount,
                TargetConditionCount = targetConditions.Count,
                ExceptionRuleCount = exceptionRules.Count(rule => rule != null && rule.HasConfiguredConditions()),
                Status = "success",
                Reason = BuildSummaryReason(result, targetConditions.Count > 0 || hasCommonScopeFilter, exceptionRules.Count(rule => rule != null && rule.HasConfiguredConditions()))
            });

            return result;
        }

        public static DataTable BuildExportTable(IEnumerable<ReviewRow> rows)
        {
            var table = new DataTable("ParameterMissingReview");
            table.Columns.Add("Item");
            table.Columns.Add("ID");
            table.Columns.Add("Name");
            table.Columns.Add("Result");
            table.Columns.Add("Content");
            table.Columns.Add("Etc");
            table.Columns.Add("Category");
            table.Columns.Add("Family");
            table.Columns.Add("Solutions");

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
                dataRow["Solutions"] = row.Solutions ?? string.Empty;
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

        private static string BuildOkMessage(ReviewResult result)
        {
            if (result == null || result.TargetElementCount <= 0)
            {
                return "검사 대상 요소가 없습니다 !!!";
            }

            return $"{result.TargetElementCount.ToString(CultureInfo.InvariantCulture)}개 요소의 데이터가 입력 되어 있습니다 !!!";
        }

        private static string BuildSummaryReason(ReviewResult result, bool hasScopeFilter, int configuredRuleCount)
        {
            if (result == null)
            {
                return "검토 결과가 없습니다.";
            }

            if (result.TargetElementCount <= 0)
            {
                return "검사 대상 요소가 없습니다 !!! / No target elements found.";
            }

            if (result.ErrorCount <= 0)
            {
                return BuildOkMessage(result);
            }

            string ruleText = configuredRuleCount > 0
                ? $" (누락 예외 필터 {configuredRuleCount.ToString(CultureInfo.InvariantCulture)}개 적용)"
                : string.Empty;
            return $"파라미터 누락 오류가 {result.ErrorCount.ToString(CultureInfo.InvariantCulture)}건 발견되었습니다{ruleText}.";
        }

        private static string BuildScopeSummary(IEnumerable<ElementParameterCondition> targetConditions,
                                                ParameterConditionCombination targetCombinationMode,
                                                string commonTargetFilterText,
                                                string commonExcludeTargetFilterText)
        {
            List<string> parts = new List<string>();

            string targetSummary = BuildConditionSummary(targetConditions, targetCombinationMode);
            if (!string.IsNullOrWhiteSpace(targetSummary))
            {
                parts.Add($"검토 대상 필터: {targetSummary}");
            }

            if (!string.IsNullOrWhiteSpace(commonTargetFilterText))
            {
                parts.Add($"공통 포함 필터: {commonTargetFilterText}");
            }

            if (!string.IsNullOrWhiteSpace(commonExcludeTargetFilterText))
            {
                parts.Add($"공통 제외 필터: {commonExcludeTargetFilterText}");
            }

            return parts.Count == 0 ? string.Empty : string.Join(" / ", parts);
        }

        private static List<string> NormalizeExtraParameterNames(IEnumerable<string> extraParameterNames)
        {
            return (extraParameterNames ?? Enumerable.Empty<string>())
                .Where(name => !string.IsNullOrWhiteSpace(name))
                .Select(name => name.Trim())
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        private static Dictionary<string, string> ReadExtraParamValues(Document doc, Element element, IEnumerable<string> extraParameterNames)
        {
            var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            if (doc == null || element == null) return result;

            foreach (string name in extraParameterNames ?? Enumerable.Empty<string>())
            {
                if (string.IsNullOrWhiteSpace(name)) continue;
                result[name] = ModelParameterExtractionService.GetElementParameterValue(doc, element, name) ?? string.Empty;
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

        private static List<string> NormalizeParameterNames(IEnumerable<string> parameterNames)
        {
            return (parameterNames ?? Enumerable.Empty<string>())
                .Where(name => !string.IsNullOrWhiteSpace(name))
                .Select(name => name.Trim())
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        private static List<MissingRule> NormalizeExceptionRules(IEnumerable<MissingRule> rules, IEnumerable<string> parameterNames)
        {
            HashSet<string> allowedNames = new HashSet<string>(NormalizeParameterNames(parameterNames), StringComparer.OrdinalIgnoreCase);
            return (rules ?? Enumerable.Empty<MissingRule>())
                .Where(rule => rule != null && rule.Enabled)
                .Select(rule =>
                {
                    MissingRule clone = rule.Clone();
                    clone.ParameterName = (clone.ParameterName ?? string.Empty).Trim();
                    return clone;
                })
                .Where(rule => allowedNames.Contains(rule.ParameterName ?? string.Empty) && rule.HasConfiguredConditions())
                .ToList();
        }

        private static List<ElementParameterCondition> GetEnabledConditions(ElementParameterUpdateSettings settings)
        {
            if (settings == null || settings.Conditions == null) return new List<ElementParameterCondition>();
            return settings.Conditions
                .Where(condition => condition != null && condition.IsConfigured() && (condition.Enabled || !string.IsNullOrWhiteSpace(condition.ParameterName)))
                .Select(condition => condition.Clone())
                .ToList();
        }

        private static List<ElementParameterCondition> GetEnabledConditions(IEnumerable<ElementParameterCondition> conditions)
        {
            return (conditions ?? Enumerable.Empty<ElementParameterCondition>())
                .Where(condition => condition != null && condition.IsConfigured() && (condition.Enabled || !string.IsNullOrWhiteSpace(condition.ParameterName)))
                .Select(condition => condition.Clone())
                .ToList();
        }

        private static bool MatchesConditions(Document doc, Element element, IEnumerable<ElementParameterCondition> conditions, ParameterConditionCombination combinationMode)
        {
            List<ElementParameterCondition> source = (conditions ?? Enumerable.Empty<ElementParameterCondition>())
                .Where(condition => condition != null)
                .ToList();

            if (source.Count == 0) return true;

            List<bool> results = source
                .Select(condition => EvaluateCondition(doc, element, condition))
                .ToList();

            return combinationMode == ParameterConditionCombination.Or
                ? results.Any(result => result)
                : results.All(result => result);
        }

        private static bool EvaluateCondition(Document doc, Element element, ElementParameterCondition condition)
        {
            if (condition == null) return true;

            string actualText;
            bool hasVirtualValue = TryGetVirtualConditionValue(doc, element, condition.ParameterName, out actualText);
            bool hasParameter = hasVirtualValue || ModelParameterExtractionService.HasElementParameter(doc, element, condition.ParameterName);
            if (!hasParameter) return false;

            if (!hasVirtualValue)
            {
                actualText = ModelParameterExtractionService.GetElementParameterValue(doc, element, condition.ParameterName);
            }

            if (condition.Operator == FilterRuleOperator.HasValue)
            {
                return !string.IsNullOrWhiteSpace(actualText);
            }

            if (condition.Operator == FilterRuleOperator.HasNoValue)
            {
                return string.IsNullOrWhiteSpace(actualText);
            }

            string expectedText = condition.Value ?? string.Empty;
            double numericActual;
            double numericExpected;
            bool hasNumericActual = TryParseNumber(actualText, out numericActual);
            bool hasNumericExpected = TryParseNumber(expectedText, out numericExpected);

            if (hasNumericActual && hasNumericExpected)
            {
                switch (condition.Operator)
                {
                    case FilterRuleOperator.Equals:
                        return Math.Abs(numericActual - numericExpected) < 0.000001d;
                    case FilterRuleOperator.NotEquals:
                        return Math.Abs(numericActual - numericExpected) >= 0.000001d;
                    case FilterRuleOperator.Greater:
                        return numericActual > numericExpected;
                    case FilterRuleOperator.GreaterOrEqual:
                        return numericActual >= numericExpected;
                    case FilterRuleOperator.Less:
                        return numericActual < numericExpected;
                    case FilterRuleOperator.LessOrEqual:
                        return numericActual <= numericExpected;
                }
            }

            string left = (actualText ?? string.Empty).Trim();
            string right = expectedText.Trim();

            switch (condition.Operator)
            {
                case FilterRuleOperator.Equals:
                    return string.Equals(left, right, StringComparison.OrdinalIgnoreCase);
                case FilterRuleOperator.NotEquals:
                    return !string.Equals(left, right, StringComparison.OrdinalIgnoreCase);
                case FilterRuleOperator.Contains:
                    return left.IndexOf(right, StringComparison.OrdinalIgnoreCase) >= 0;
                case FilterRuleOperator.NotContains:
                    return left.IndexOf(right, StringComparison.OrdinalIgnoreCase) < 0;
                case FilterRuleOperator.BeginsWith:
                    return left.StartsWith(right, StringComparison.OrdinalIgnoreCase);
                case FilterRuleOperator.NotBeginsWith:
                    return !left.StartsWith(right, StringComparison.OrdinalIgnoreCase);
                case FilterRuleOperator.EndsWith:
                    return left.EndsWith(right, StringComparison.OrdinalIgnoreCase);
                case FilterRuleOperator.NotEndsWith:
                    return !left.EndsWith(right, StringComparison.OrdinalIgnoreCase);
                case FilterRuleOperator.Greater:
                    return string.Compare(left, right, StringComparison.OrdinalIgnoreCase) > 0;
                case FilterRuleOperator.GreaterOrEqual:
                    return string.Compare(left, right, StringComparison.OrdinalIgnoreCase) >= 0;
                case FilterRuleOperator.Less:
                    return string.Compare(left, right, StringComparison.OrdinalIgnoreCase) < 0;
                case FilterRuleOperator.LessOrEqual:
                    return string.Compare(left, right, StringComparison.OrdinalIgnoreCase) <= 0;
                default:
                    return false;
            }
        }

        private static bool TryGetVirtualConditionValue(Document doc, Element element, string parameterName, out string value)
        {
            value = string.Empty;
            string token = NormalizeConditionParameterToken(parameterName);
            if (string.IsNullOrWhiteSpace(token)) return false;

            if (IsAnyConditionToken(token, "category", "categoryname", "cat", "카테고리", "분류"))
            {
                value = ModelParameterExtractionService.GetElementCategoryName(element) ?? string.Empty;
                return true;
            }

            if (IsAnyConditionToken(token, "family", "familyname", "fam", "패밀리", "패밀리명"))
            {
                value = ModelParameterExtractionService.GetElementFamilyName(doc, element) ?? string.Empty;
                return true;
            }

            return false;
        }

        private static bool IsAnyConditionToken(string token, params string[] candidates)
        {
            return candidates != null
                   && candidates.Any(candidate => string.Equals(token, NormalizeConditionParameterToken(candidate), StringComparison.OrdinalIgnoreCase));
        }

        private static string NormalizeConditionParameterToken(string value)
        {
            if (string.IsNullOrWhiteSpace(value)) return string.Empty;

            var builder = new StringBuilder(value.Length);
            foreach (char ch in value.Trim())
            {
                if (char.IsWhiteSpace(ch) || ch == '_' || ch == '-') continue;
                builder.Append(char.ToLowerInvariant(ch));
            }

            return builder.ToString();
        }

        private static bool ShouldIgnoreMissing(Document doc, Element element, string parameterName, IEnumerable<MissingRule> rules, out MissingRule matchedRule)
        {
            foreach (MissingRule rule in rules ?? Enumerable.Empty<MissingRule>())
            {
                if (rule == null) continue;
                if (!string.Equals(rule.ParameterName, parameterName, StringComparison.OrdinalIgnoreCase)) continue;

                List<ElementParameterCondition> conditions = GetEnabledConditions(rule.Conditions);
                if (conditions.Count == 0) continue;

                if (MatchesConditions(doc, element, conditions, rule.CombinationMode))
                {
                    matchedRule = rule;
                    return true;
                }
            }

            matchedRule = null;
            return false;
        }

        private static ParameterReviewState GetParameterReviewState(Document doc, Element element, string parameterName)
        {
            if (!ModelParameterExtractionService.HasElementParameter(doc, element, parameterName))
            {
                return ParameterReviewState.ParameterNotFound;
            }

            string value = ModelParameterExtractionService.GetElementParameterValue(doc, element, parameterName);
            return string.IsNullOrWhiteSpace(value)
                ? ParameterReviewState.EmptyValue
                : ParameterReviewState.HasValue;
        }

        private static string BuildItemText(int caseCount)
        {
            int safeCount = Math.Max(0, caseCount);
            string countText = safeCount.ToString(CultureInfo.InvariantCulture);
            return $"속성누락검토 오류 ({countText}건)";
        }

        private static string BuildContentMessage(string parameterName, ParameterReviewState reviewState)
        {
            string safeParameterName = string.IsNullOrWhiteSpace(parameterName) ? string.Empty : parameterName.Trim();
            if (reviewState == ParameterReviewState.ParameterNotFound)
            {
                return $"[Parameter] : [{safeParameterName}] 파라미터가 존재 하지 않습니다.";
            }

            return $"[Parameter]: [{safeParameterName}] 값이 누락입니다.";
        }

        private static string BuildSolutionsMessage(ParameterReviewState reviewState)
        {
            if (reviewState == ParameterReviewState.ParameterNotFound)
            {
                return string.Empty;
            }

            return "속성 값을 입력해주세요.";
        }

        private static string BuildRuleSummary(IEnumerable<MissingRule> rules, string parameterName)
        {
            List<string> summaries = (rules ?? Enumerable.Empty<MissingRule>())
                .Where(rule => rule != null && string.Equals(rule.ParameterName, parameterName, StringComparison.OrdinalIgnoreCase))
                .Select(rule => rule.BuildSummary())
                .Where(text => !string.IsNullOrWhiteSpace(text))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();

            if (summaries.Count == 0) return string.Empty;
            if (summaries.Count == 1) return summaries[0];
            return string.Join(" OR ", summaries);
        }

        private static string BuildConditionSummary(IEnumerable<ElementParameterCondition> conditions, ParameterConditionCombination combinationMode)
        {
            List<string> items = (conditions ?? Enumerable.Empty<ElementParameterCondition>())
                .Where(condition => condition != null && condition.IsConfigured())
                .Select(condition => condition.ToString())
                .Where(text => !string.IsNullOrWhiteSpace(text))
                .ToList();

            if (items.Count == 0) return string.Empty;

            string joiner = combinationMode == ParameterConditionCombination.Or ? " OR " : " AND ";
            return string.Join(joiner, items);
        }

        private static TypeInfo ResolveTypeInfo(Document doc, Element element)
        {
            return new TypeInfo
            {
                Category = ModelParameterExtractionService.GetElementCategoryName(element),
                Family = ModelParameterExtractionService.GetElementFamilyName(doc, element),
                TypeName = ModelParameterExtractionService.GetElementTypeName(doc, element)
            };
        }

        private static bool TryParseNumber(string text, out double value)
        {
            NumberStyles styles = NumberStyles.Float | NumberStyles.AllowThousands;
            if (double.TryParse(text, styles, CultureInfo.InvariantCulture, out value))
            {
                return true;
            }

            return double.TryParse(text, styles, CultureInfo.CurrentCulture, out value);
        }
    }
}
