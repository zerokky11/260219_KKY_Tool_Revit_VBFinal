using System.Collections.Generic;
using System.Linq;

namespace KKY_Tool_Revit.Models
{
    public class BatchCleanSettings
    {
        public List<string> FilePaths { get; set; } = new List<string>();
        public string OutputFolder { get; set; }
        public string Target3DViewName { get; set; }
        public List<ViewParameterAssignment> ViewParameters { get; set; } = new List<ViewParameterAssignment>();

        public bool UseFilter { get; set; }
        public bool ApplyFilterInitially { get; set; }
        public bool AutoEnableFilterIfEmpty { get; set; }
        public ViewFilterProfile FilterProfile { get; set; }
        public ParameterConditionCombination VisibilityRuleCombinationMode { get; set; } = ParameterConditionCombination.Or;
        public List<VisibilitySubCategoryRule> VisibilitySubCategoryRules { get; set; } = new List<VisibilitySubCategoryRule>();
        public bool? ShowImportedCategoriesInView { get; set; }
        public bool? ShowImportsInFamilies { get; set; }

        public ElementParameterUpdateSettings ElementParameterUpdate { get; set; } = new ElementParameterUpdateSettings();

        public BatchCleanSettings Clone()
        {
            return new BatchCleanSettings
            {
                FilePaths = FilePaths != null ? new List<string>(FilePaths) : new List<string>(),
                OutputFolder = OutputFolder,
                Target3DViewName = Target3DViewName,
                ViewParameters = ViewParameters != null
                    ? ViewParameters.Where(x => x != null).Select(x => new ViewParameterAssignment
                    {
                        Enabled = x.Enabled,
                        ParameterName = x.ParameterName,
                        ParameterValue = x.ParameterValue
                    }).ToList()
                    : new List<ViewParameterAssignment>(),
                UseFilter = UseFilter,
                ApplyFilterInitially = ApplyFilterInitially,
                AutoEnableFilterIfEmpty = AutoEnableFilterIfEmpty,
                FilterProfile = FilterProfile != null ? FilterProfile.Clone() : null,
                VisibilityRuleCombinationMode = VisibilityRuleCombinationMode,
                VisibilitySubCategoryRules = VisibilitySubCategoryRules != null
                    ? VisibilitySubCategoryRules.Where(x => x != null).Select(x => x.Clone()).ToList()
                    : new List<VisibilitySubCategoryRule>(),
                ShowImportedCategoriesInView = ShowImportedCategoriesInView,
                ShowImportsInFamilies = ShowImportsInFamilies,
                ElementParameterUpdate = ElementParameterUpdate != null ? ElementParameterUpdate.Clone() : new ElementParameterUpdateSettings()
            };
        }
    }
}
