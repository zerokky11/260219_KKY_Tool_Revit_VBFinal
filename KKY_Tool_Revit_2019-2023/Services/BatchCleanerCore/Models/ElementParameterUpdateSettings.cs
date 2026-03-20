using System;
using System.Collections.Generic;
using System.Linq;

namespace KKY_Tool_Revit.Models
{
    public enum ParameterConditionCombination
    {
        And,
        Or
    }

    public class ElementParameterCondition
    {
        public bool Enabled { get; set; }
        public string ParameterName { get; set; }
        public FilterRuleOperator Operator { get; set; }
        public string Value { get; set; }

        public bool IsConfigured()
        {
            if (string.IsNullOrWhiteSpace(ParameterName)) return false;

            switch (Operator)
            {
                case FilterRuleOperator.HasValue:
                case FilterRuleOperator.HasNoValue:
                    return true;
                default:
                    return Value != null;
            }
        }

        public ElementParameterCondition Clone()
        {
            return new ElementParameterCondition
            {
                Enabled = Enabled,
                ParameterName = ParameterName,
                Operator = Operator,
                Value = Value
            };
        }

        public override string ToString()
        {
            if (!IsConfigured()) return string.Empty;

            switch (Operator)
            {
                case FilterRuleOperator.HasValue:
                case FilterRuleOperator.HasNoValue:
                    return ParameterName + " " + Operator;
                default:
                    return ParameterName + " " + Operator + " " + (Value ?? string.Empty);
            }
        }
    }

    public class ElementParameterAssignment
    {
        public bool Enabled { get; set; }
        public string ParameterName { get; set; }
        public string Value { get; set; }

        public bool IsConfigured()
        {
            return !string.IsNullOrWhiteSpace(ParameterName);
        }

        public ElementParameterAssignment Clone()
        {
            return new ElementParameterAssignment
            {
                Enabled = Enabled,
                ParameterName = ParameterName,
                Value = Value
            };
        }

        public override string ToString()
        {
            return IsConfigured() ? (ParameterName + " = " + (Value ?? string.Empty)) : string.Empty;
        }
    }

    public class ElementParameterUpdateSettings
    {
        public bool Enabled { get; set; }
        public ParameterConditionCombination CombinationMode { get; set; } = ParameterConditionCombination.And;
        public List<ElementParameterCondition> Conditions { get; set; } = new List<ElementParameterCondition>();
        public List<ElementParameterAssignment> Assignments { get; set; } = new List<ElementParameterAssignment>();

        public bool IsConfigured()
        {
            return Conditions.Any(x => x != null && x.IsConfigured())
                   && Assignments.Any(x => x != null && x.IsConfigured());
        }

        public ElementParameterUpdateSettings Clone()
        {
            return new ElementParameterUpdateSettings
            {
                Enabled = Enabled,
                CombinationMode = CombinationMode,
                Conditions = Conditions.Where(x => x != null).Select(x => x.Clone()).ToList(),
                Assignments = Assignments.Where(x => x != null).Select(x => x.Clone()).ToList()
            };
        }

        public string BuildSummary()
        {
            var conditions = Conditions
                .Where(x => x != null && x.IsConfigured())
                .Select(x => x.ToString())
                .Where(x => !string.IsNullOrWhiteSpace(x))
                .ToList();

            var assignments = Assignments
                .Where(x => x != null && x.IsConfigured())
                .Select(x => x.ToString())
                .Where(x => !string.IsNullOrWhiteSpace(x))
                .ToList();

            if (conditions.Count == 0 && assignments.Count == 0)
            {
                return "설정 없음";
            }

            string joiner = CombinationMode == ParameterConditionCombination.Or ? " OR " : " AND ";
            string conditionText = conditions.Count == 0 ? "조건 없음" : string.Join(joiner, conditions);
            string assignmentText = assignments.Count == 0 ? "입력 없음" : string.Join(" / ", assignments);
            return "조건: " + conditionText + Environment.NewLine + "입력: " + assignmentText;
        }
    }
}
