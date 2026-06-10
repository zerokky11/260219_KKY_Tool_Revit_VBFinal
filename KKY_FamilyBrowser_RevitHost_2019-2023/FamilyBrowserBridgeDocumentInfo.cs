public sealed class FamilyBrowserBridgeDocumentInfo
{
	public string DocumentId { get; set; }

	public string Title { get; set; }

	public string Path { get; set; }

	public string CentralPath { get; set; }

	public bool IsActive { get; set; }

	public bool IsWorkshared { get; set; }

	public FamilyBrowserBridgeDocumentInfo()
	{
		DocumentId = string.Empty;
		Title = string.Empty;
		Path = string.Empty;
		CentralPath = string.Empty;
		IsActive = false;
		IsWorkshared = false;
	}
}
