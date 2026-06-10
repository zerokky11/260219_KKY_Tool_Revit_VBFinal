using System.Collections.Generic;

public class StandardLibrarySnapshot
{
	public int SnapshotSchemaVersion { get; set; }

	public string SourceId { get; set; }

	public string DisplayName { get; set; }

	public string SourceKind { get; set; }

	public string Locator { get; set; }

	public string ResolvedPath { get; set; }

	public string SnapshotMode { get; set; }

	public string SourceFileLastWriteUtc { get; set; }

	public long SourceFileLength { get; set; }

	public string CapturedAtUtc { get; set; }

	public string RevitVersion { get; set; }

	public StandardLibrarySnapshotSummary Summary { get; set; }

	public List<StandardLoadableFamilySnapshotItem> LoadableFamilies { get; set; }

	public List<StandardSystemTypeSnapshotItem> SystemTypes { get; set; }

	public StandardLibrarySnapshot()
	{
		SnapshotSchemaVersion = 5;
		SourceId = string.Empty;
		DisplayName = string.Empty;
		SourceKind = string.Empty;
		Locator = string.Empty;
		ResolvedPath = string.Empty;
		SnapshotMode = string.Empty;
		SourceFileLastWriteUtc = string.Empty;
		CapturedAtUtc = string.Empty;
		RevitVersion = string.Empty;
		Summary = new StandardLibrarySnapshotSummary();
		LoadableFamilies = new List<StandardLoadableFamilySnapshotItem>();
		SystemTypes = new List<StandardSystemTypeSnapshotItem>();
	}
}
