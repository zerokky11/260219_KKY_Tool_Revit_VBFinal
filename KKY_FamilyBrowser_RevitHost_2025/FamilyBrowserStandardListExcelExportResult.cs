public sealed class FamilyBrowserStandardListExcelExportResult
{
	public string OutputPath { get; set; }

	public int RowCount { get; set; }

	public string SheetName { get; set; }

	public FamilyBrowserStandardListExcelExportResult()
	{
		OutputPath = string.Empty;
		RowCount = 0;
		SheetName = "StandardList";
	}
}
