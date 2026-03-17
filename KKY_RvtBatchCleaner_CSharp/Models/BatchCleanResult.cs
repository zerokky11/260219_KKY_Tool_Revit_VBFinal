namespace KKY_Tool_Revit.Models
{
    public class BatchCleanResult
    {
        public string SourcePath { get; set; }
        public string OutputPath { get; set; }
        public bool Success { get; set; }
        public string Message { get; set; }
    }
}
