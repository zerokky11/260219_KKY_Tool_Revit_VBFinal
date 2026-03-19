using Autodesk.Revit.DB;

namespace KKY_Tool_Revit.Models
{
    public class PreparedDocumentEntry
    {
        public string SourcePath { get; set; }
        public string OutputPath { get; set; }
        public Document Document { get; set; }
        public ElementId TargetViewId { get; set; }
        public ElementId KeptFilterId { get; set; }
    }
}
