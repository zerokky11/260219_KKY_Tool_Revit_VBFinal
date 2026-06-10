namespace KKY_FamilyBrowser_RevitHost_2025;

internal class DocumentFamilyDiagnosticsItem
{
	public string FamilyName { get; set; }

	public string CategoryName { get; set; }

	public bool IsEditable { get; set; }

	public bool IsInPlace { get; set; }

	public bool IsShared { get; set; }

	public int TypeCount { get; set; }

	public string UniqueId { get; set; }

	public DocumentFamilyDiagnosticsItem()
	{
		FamilyName = string.Empty;
		CategoryName = string.Empty;
		UniqueId = string.Empty;
	}
}
