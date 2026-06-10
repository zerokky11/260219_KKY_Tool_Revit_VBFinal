public class FamilyBrowserRequestStoreInfo
{
	public string Mode { get; set; }

	public string StoreLocation { get; set; }

	public string Endpoint { get; set; }

	public string DisplayName { get; set; }

	public string Detail { get; set; }

	public bool IsShared { get; set; }

	public bool IsFileBacked { get; set; }

	public bool UsesConnectorQueue { get; set; }

	public bool RequiresConnectorSync { get; set; }

	public FamilyBrowserRequestStoreInfo()
	{
		Mode = string.Empty;
		StoreLocation = string.Empty;
		Endpoint = string.Empty;
		DisplayName = string.Empty;
		Detail = string.Empty;
		IsShared = false;
		IsFileBacked = true;
		UsesConnectorQueue = false;
		RequiresConnectorSync = false;
	}
}
