using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using KKY_Tool_Revit.Models;

namespace KKY_Tool_Revit.Services
{
    public static class VisibilitySubCategoryRuleProfileService
    {
        public static void SaveToXml(ParameterConditionCombination combinationMode, IEnumerable<VisibilitySubCategoryRule> rules, string filePath)
        {
            SaveToXml(combinationMode, rules, null, null, filePath);
        }

        public static void SaveToXml(ParameterConditionCombination combinationMode,
            IEnumerable<VisibilitySubCategoryRule> rules,
            bool? showImportedCategoriesInView,
            bool? showImportsInFamilies,
            string filePath)
        {
            if (string.IsNullOrWhiteSpace(filePath)) throw new ArgumentNullException(nameof(filePath));

            var root = new XElement("VisibilityRuleProfile",
                new XAttribute("CombinationMode", combinationMode.ToString()));

            if (showImportedCategoriesInView.HasValue)
            {
                root.Add(new XAttribute("ShowImportedCategoriesInView", showImportedCategoriesInView.Value));
            }

            if (showImportsInFamilies.HasValue)
            {
                root.Add(new XAttribute("ShowImportsInFamilies", showImportsInFamilies.Value));
            }

            root.Add((rules ?? Enumerable.Empty<VisibilitySubCategoryRule>())
                .Where(x => x != null)
                .Select(CreateRuleElement));

            new XDocument(root).Save(filePath);
        }

        public static VisibilityRuleProfileSnapshot LoadFromXml(string filePath)
        {
            if (string.IsNullOrWhiteSpace(filePath)) throw new ArgumentNullException(nameof(filePath));

            XDocument doc = XDocument.Load(filePath);
            XElement root = doc.Root ?? throw new InvalidOperationException("VV 규칙 파일 형식이 올바르지 않습니다.");

            ParameterConditionCombination combinationMode = ParameterConditionCombination.Or;
            Enum.TryParse((string)root.Attribute("CombinationMode"), true, out combinationMode);

            var rules = root.Elements("Rule")
                .Select(ParseRuleElement)
                .Where(x => x != null)
                .ToList();

            return new VisibilityRuleProfileSnapshot
            {
                CombinationMode = combinationMode,
                Rules = rules,
                ShowImportedCategoriesInView = ParseNullableBool((string)root.Attribute("ShowImportedCategoriesInView")),
                ShowImportsInFamilies = ParseNullableBool((string)root.Attribute("ShowImportsInFamilies"))
            };
        }

        private static XElement CreateRuleElement(VisibilitySubCategoryRule rule)
        {
            return new XElement("Rule",
                new XAttribute("Enabled", rule.Enabled),
                new XAttribute("Operator", rule.Operator.ToString()),
                new XElement("SubCategoryText", rule.SubCategoryText ?? string.Empty),
                new XElement("ParentCategories",
                    rule.GetNormalizedParentCategoryNames().Select(x => new XElement("Category", x))));
        }

        private static VisibilitySubCategoryRule ParseRuleElement(XElement element)
        {
            if (element == null) return null;

            FilterRuleOperator op = FilterRuleOperator.Contains;
            Enum.TryParse((string)element.Attribute("Operator"), true, out op);

            return new VisibilitySubCategoryRule
            {
                Enabled = ParseBool((string)element.Attribute("Enabled"), true),
                Operator = op,
                SubCategoryText = (string)element.Element("SubCategoryText") ?? string.Empty,
                ParentCategoryNames = element.Element("ParentCategories")?
                    .Elements("Category")
                    .Select(x => (x.Value ?? string.Empty).Trim())
                    .Where(x => !string.IsNullOrWhiteSpace(x))
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .ToList() ?? new List<string>()
            };
        }

        private static bool ParseBool(string text, bool defaultValue)
        {
            bool parsed;
            return bool.TryParse(text, out parsed) ? parsed : defaultValue;
        }

        private static bool? ParseNullableBool(string text)
        {
            bool parsed;
            return bool.TryParse(text, out parsed) ? parsed : (bool?)null;
        }

        public class VisibilityRuleProfileSnapshot
        {
            public ParameterConditionCombination CombinationMode { get; set; } = ParameterConditionCombination.Or;
            public List<VisibilitySubCategoryRule> Rules { get; set; } = new List<VisibilitySubCategoryRule>();
            public bool? ShowImportedCategoriesInView { get; set; }
            public bool? ShowImportsInFamilies { get; set; }
        }
    }
}
