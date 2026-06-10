public sealed class FamilyBrowserPermissionExcelSettings
{
	public bool Enabled { get; set; }

	public string Path { get; set; }

	public string SheetName { get; set; }

	public string LastUpdatedUtc { get; set; }

	public string LastUpdatedBy { get; set; }

	public FamilyBrowserPermissionExcelSettings()
	{
		Enabled = false;
		Path = string.Empty;
		SheetName = string.Empty;
		LastUpdatedUtc = string.Empty;
		LastUpdatedBy = string.Empty;
	}

	public static FamilyBrowserPermissionExcelSettings CreateDefault()
	{
		return new FamilyBrowserPermissionExcelSettings();
	}
}
