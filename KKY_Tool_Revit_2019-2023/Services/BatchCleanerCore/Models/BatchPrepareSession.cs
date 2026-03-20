using System.Collections.Generic;

namespace KKY_Tool_Revit.Models
{
    public class ModelObjectCountComparison
    {
        public string FileName { get; set; }
        public string SourcePath { get; set; }
        public string OutputPath { get; set; }
        public int? BeforeCount { get; set; }
        public int? AfterCount { get; set; }
        public string Status { get; set; }
        public string Note { get; set; }
    }

    public class BatchPrepareSession
    {
        public List<PreparedDocumentEntry> PreparedDocuments { get; set; } = new List<PreparedDocumentEntry>();
        public List<BatchCleanResult> Results { get; set; } = new List<BatchCleanResult>();
        public List<string> CleanedOutputPaths { get; set; } = new List<string>();
        public List<ModelObjectCountComparison> CleanCountComparisons { get; set; } = new List<ModelObjectCountComparison>();
        public List<ModelObjectCountComparison> PurgeCountComparisons { get; set; } = new List<ModelObjectCountComparison>();
        public string OutputFolder { get; set; }
        public string DesignOptionAuditCsvPath { get; set; }
        public string VerificationCsvPath { get; set; }
        public string ExtractionCsvPath { get; set; }
        public string PurgeCountComparisonXlsxPath { get; set; }
    }
}
