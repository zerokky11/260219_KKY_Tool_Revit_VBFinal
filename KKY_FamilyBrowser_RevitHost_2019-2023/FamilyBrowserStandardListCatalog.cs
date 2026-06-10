using System.Collections.Generic;

public sealed class FamilyBrowserStandardListCatalog
{
	public string SourcePath { get; set; }

	public string SheetName { get; set; }

	public bool ExplicitPath { get; set; }

	public bool Exists { get; set; }

	public int RowCount { get; set; }

	public string LastLoadedUtc { get; set; }

	public string LastError { get; set; }

	public List<FamilyBrowserStandardListEntry> Entries { get; set; }

	public string BaselineCreatedAtUtc { get; set; }

	public string BaselineCreatedBy { get; set; }

	public string BaselineSourceSnapshotPath { get; set; }

	public int BaselineSystemExclusionVersion { get; set; }

	public List<FamilyBrowserStandardListEntry> BaselineExcludedLoadableFamilies { get; set; }

	public List<FamilyBrowserStandardListEntry> BaselineExcludedSystemTypes { get; set; }

	public FamilyBrowserStandardListCatalog()
	{
		SourcePath = string.Empty;
		SheetName = string.Empty;
		LastLoadedUtc = string.Empty;
		LastError = string.Empty;
		Entries = new List<FamilyBrowserStandardListEntry>();
		BaselineCreatedAtUtc = string.Empty;
		BaselineCreatedBy = string.Empty;
		BaselineSourceSnapshotPath = string.Empty;
		BaselineExcludedLoadableFamilies = new List<FamilyBrowserStandardListEntry>();
		BaselineExcludedSystemTypes = new List<FamilyBrowserStandardListEntry>();
	}
}
