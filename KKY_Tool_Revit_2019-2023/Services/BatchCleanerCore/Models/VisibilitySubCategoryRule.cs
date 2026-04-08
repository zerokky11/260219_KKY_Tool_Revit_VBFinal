using System;
using System.Collections.Generic;
using System.Linq;

namespace KKY_Tool_Revit.Models
{
    public class VisibilitySubCategoryRule
    {
        public bool Enabled { get; set; }
        public List<string> ParentCategoryNames { get; set; } = new List<string>();
        public FilterRuleOperator Operator { get; set; } = FilterRuleOperator.Contains;
        public string SubCategoryText { get; set; }

        public IEnumerable<string> GetNormalizedParentCategoryNames()
        {
            return (ParentCategoryNames ?? new List<string>())
                .Where(x => !string.IsNullOrWhiteSpace(x))
                .Select(x => x.Trim())
                .Distinct(StringComparer.OrdinalIgnoreCase);
        }

        public bool IsConfigured()
        {
            if (!GetNormalizedParentCategoryNames().Any())
            {
                return false;
            }

            switch (Operator)
            {
                case FilterRuleOperator.HasValue:
                case FilterRuleOperator.HasNoValue:
                    return true;
                default:
                    return !string.IsNullOrWhiteSpace(SubCategoryText);
            }
        }

        public VisibilitySubCategoryRule Clone()
        {
            return new VisibilitySubCategoryRule
            {
                Enabled = Enabled,
                ParentCategoryNames = GetNormalizedParentCategoryNames().ToList(),
                Operator = Operator,
                SubCategoryText = SubCategoryText
            };
        }

        public override string ToString()
        {
            if (!IsConfigured()) return string.Empty;

            switch (Operator)
            {
                case FilterRuleOperator.HasValue:
                case FilterRuleOperator.HasNoValue:
                    return string.Join(", ", GetNormalizedParentCategoryNames()) + " / " + Operator;
                default:
                    return string.Join(", ", GetNormalizedParentCategoryNames()) + " / " + Operator + " / " + (SubCategoryText ?? string.Empty);
            }
        }
    }
}
