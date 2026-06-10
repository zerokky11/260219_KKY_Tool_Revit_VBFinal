public sealed class FamilyBrowserPermissionExcelExportResult
{
	public string OutputPath { get; set; }

	public string SourceFolder { get; set; }

	public int RowCount { get; set; }

	public int SkippedBackupCount { get; set; }

	public string SheetName { get; set; }

	public FamilyBrowserPermissionExcelExportResult()
	{
		OutputPath = string.Empty;
		SourceFolder = string.Empty;
		RowCount = 0;
		SkippedBackupCount = 0;
		SheetName = "Policy";
	}
}
