using System.Collections.Generic;

namespace KKY_Tool_Revit.Models
{
    public class BatchPrepareSession
    {
        public List<PreparedDocumentEntry> PreparedDocuments { get; set; } = new List<PreparedDocumentEntry>();
        public List<BatchCleanResult> Results { get; set; } = new List<BatchCleanResult>();
        public List<string> CleanedOutputPaths { get; set; } = new List<string>();
        public string OutputFolder { get; set; }
        public string DesignOptionAuditCsvPath { get; set; }
        public string VerificationCsvPath { get; set; }
    }
}
