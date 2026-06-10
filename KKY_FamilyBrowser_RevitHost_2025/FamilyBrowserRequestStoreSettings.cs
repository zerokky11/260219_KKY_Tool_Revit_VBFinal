public sealed class FamilyBrowserRequestStoreSettings
{
	public string Mode { get; set; }

	public string Path { get; set; }

	public string Endpoint { get; set; }

	public string LastUpdatedUtc { get; set; }

	public string LastUpdatedBy { get; set; }

	public FamilyBrowserRequestStoreSettings()
	{
		Mode = "Local";
		Path = string.Empty;
		Endpoint = string.Empty;
		LastUpdatedUtc = string.Empty;
		LastUpdatedBy = string.Empty;
	}

	public static FamilyBrowserRequestStoreSettings CreateDefault()
	{
		return new FamilyBrowserRequestStoreSettings
		{
			Mode = "Local"
		};
	}
}
