public sealed class FamilyBrowserPermissionExcelStatus
{
	public bool Enabled { get; set; }

	public string Path { get; set; }

	public string SheetName { get; set; }

	public bool Exists { get; set; }

	public int RowCount { get; set; }

	public string LastLoadedUtc { get; set; }

	public string LastError { get; set; }

	public FamilyBrowserPermissionExcelStatus()
	{
		Enabled = false;
		Path = string.Empty;
		SheetName = string.Empty;
		Exists = false;
		RowCount = 0;
		LastLoadedUtc = string.Empty;
		LastError = string.Empty;
	}
}
