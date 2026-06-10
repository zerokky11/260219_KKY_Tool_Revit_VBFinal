public sealed class ProjectComparisonReviewExcelExportResult
{
	public string OutputPath { get; set; }

	public int RowCount { get; set; }

	public string SheetName { get; set; }

	public ProjectComparisonReviewExcelExportResult()
	{
		OutputPath = string.Empty;
		SheetName = "Review";
	}
}
