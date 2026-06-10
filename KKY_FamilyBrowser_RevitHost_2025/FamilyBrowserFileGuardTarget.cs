public sealed class FamilyBrowserFileGuardTarget
{
	public bool Enabled { get; set; }

	public string FileName { get; set; }

	public string CentralPath { get; set; }

	public string RelativePath { get; set; }

	public bool BlockFamilyLoadAndEdit { get; set; }

	public bool BlockTypeChanges { get; set; }

	public string LastUpdatedUtc { get; set; }

	public string LastUpdatedBy { get; set; }

	public FamilyBrowserFileGuardTarget()
	{
		Enabled = true;
		FileName = string.Empty;
		CentralPath = string.Empty;
		RelativePath = string.Empty;
		BlockFamilyLoadAndEdit = true;
		BlockTypeChanges = true;
		LastUpdatedUtc = string.Empty;
		LastUpdatedBy = string.Empty;
	}
}
