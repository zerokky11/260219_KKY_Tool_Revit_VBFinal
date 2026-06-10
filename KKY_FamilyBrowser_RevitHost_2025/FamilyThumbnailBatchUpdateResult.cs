using System.Collections.Generic;

public class FamilyThumbnailBatchUpdateResult
{
	public int SuccessCount { get; set; }

	public int FailedCount { get; set; }

	public int SkippedCount { get; set; }

	public string OutputFolder { get; set; }

	public string DiagnosticReportPath { get; set; }

	public List<FamilyThumbnailBatchUpdateItem> Items { get; set; }

	public List<FamilyThumbnailAutoConfirmedDialogRecord> AutoConfirmedDialogs { get; set; }

	public FamilyThumbnailBatchUpdateResult()
	{
		OutputFolder = string.Empty;
		DiagnosticReportPath = string.Empty;
		Items = new List<FamilyThumbnailBatchUpdateItem>();
		AutoConfirmedDialogs = new List<FamilyThumbnailAutoConfirmedDialogRecord>();
	}
}
