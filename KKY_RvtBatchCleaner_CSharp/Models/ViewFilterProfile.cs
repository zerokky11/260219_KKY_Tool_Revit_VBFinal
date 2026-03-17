using System;
using System.Collections.Generic;
using System.Linq;

namespace KKY_Tool_Revit.Models
{
    public enum FilterRuleOperator
    {
        Equals,
        NotEquals,
        Contains,
        NotContains,
        BeginsWith,
        NotBeginsWith,
        EndsWith,
        NotEndsWith,
        Greater,
        GreaterOrEqual,
        Less,
        LessOrEqual,
        HasValue,
        HasNoValue
    }

    public class ViewFilterProfile
    {
        public string FilterName { get; set; }
        public string CategoriesCsv { get; set; }
        public string ParameterToken { get; set; }
        public FilterRuleOperator Operator { get; set; }
        public string RuleValue { get; set; }
        public string FilterDefinitionXml { get; set; }
        public string StructureSummary { get; set; }

        public bool HasSerializedDefinition
        {
            get { return !string.IsNullOrWhiteSpace(FilterDefinitionXml); }
        }

        public List<string> GetCategoryTokens()
        {
            return (CategoriesCsv ?? string.Empty)
                .Split(new[] { ',', ';', '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries)
                .Select(x => x.Trim())
                .Where(x => !string.IsNullOrWhiteSpace(x))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        public bool IsConfigured()
        {
            return !string.IsNullOrWhiteSpace(FilterName)
                   && !string.IsNullOrWhiteSpace(CategoriesCsv)
                   && (HasSerializedDefinition || (!string.IsNullOrWhiteSpace(ParameterToken) && RuleValue != null));
        }

        public ViewFilterProfile Clone()
        {
            return new ViewFilterProfile
            {
                FilterName = FilterName,
                CategoriesCsv = CategoriesCsv,
                ParameterToken = ParameterToken,
                Operator = Operator,
                RuleValue = RuleValue,
                FilterDefinitionXml = FilterDefinitionXml,
                StructureSummary = StructureSummary
            };
        }
    }
}
