public sealed class FamilyBrowserFileGuardTarget
{
	public bool Enabled { get; set; }

	public string FileName { get; set; }

	public string CentralPath { get; set; }

	public string RelativePath { get; set; }

	public string Discipline { get; set; }

	public bool BlockFamilyLoadAndEdit { get; set; }

	public bool BlockTypeChanges { get; set; }

	public bool BlockNestedOnlyStandalonePlacement { get; set; }

	public bool TrackElementChanges { get; set; }

	public bool TrackElementChangesConfigured { get; set; }

	public string LastUpdatedUtc { get; set; }

	public string LastUpdatedBy { get; set; }

	public FamilyBrowserFileGuardTarget()
	{
		Enabled = true;
		FileName = string.Empty;
		CentralPath = string.Empty;
		RelativePath = string.Empty;
		Discipline = string.Empty;
		BlockFamilyLoadAndEdit = true;
		BlockTypeChanges = true;
		BlockNestedOnlyStandalonePlacement = false;
		TrackElementChanges = true;
		TrackElementChangesConfigured = false;
		LastUpdatedUtc = string.Empty;
		LastUpdatedBy = string.Empty;
	}
}
