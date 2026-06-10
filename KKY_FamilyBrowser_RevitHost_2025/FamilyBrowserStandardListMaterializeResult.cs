public sealed class FamilyBrowserStandardListMaterializeResult
{
	public string SourcePath { get; set; }

	public string OutputPath { get; set; }

	public string SheetName { get; set; }

	public int RowCount { get; set; }

	public FamilyBrowserStandardListMaterializeResult()
	{
		SourcePath = string.Empty;
		OutputPath = string.Empty;
		SheetName = string.Empty;
	}
}
