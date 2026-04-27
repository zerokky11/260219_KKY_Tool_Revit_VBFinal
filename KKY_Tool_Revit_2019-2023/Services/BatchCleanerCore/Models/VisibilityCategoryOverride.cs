using System;

namespace KKY_Tool_Revit.Models
{
    public class VisibilityCategoryOverride
    {
        public string CategoryName { get; set; }
        public bool Visible { get; set; } = true;

        public string GetNormalizedCategoryName()
        {
            return (CategoryName ?? string.Empty).Trim();
        }

        public bool IsConfigured()
        {
            return !string.IsNullOrWhiteSpace(GetNormalizedCategoryName());
        }

        public VisibilityCategoryOverride Clone()
        {
            return new VisibilityCategoryOverride
            {
                CategoryName = CategoryName,
                Visible = Visible
            };
        }
    }
}
