public class StandardLibraryRegistrationRecord
{
	public string SourceId { get; set; }

	public string DisplayName { get; set; }

	public string SourceKind { get; set; }

	public string Locator { get; set; }

	public string ResolvedPath { get; set; }

	public string SnapshotMode { get; set; }

	public string SourceFileLastWriteUtc { get; set; }

	public long SourceFileLength { get; set; }

	public string RegisteredAtUtc { get; set; }

	public string RegisteredBy { get; set; }

	public string LastSnapshotAtUtc { get; set; }

	public string LastSnapshotPath { get; set; }

	public string RevitVersion { get; set; }

	public StandardLibrarySnapshotSummary Summary { get; set; }

	public StandardLibraryRegistrationRecord()
	{
		SourceId = string.Empty;
		DisplayName = string.Empty;
		SourceKind = string.Empty;
		Locator = string.Empty;
		ResolvedPath = string.Empty;
		SnapshotMode = string.Empty;
		SourceFileLastWriteUtc = string.Empty;
		RegisteredAtUtc = string.Empty;
		RegisteredBy = string.Empty;
		LastSnapshotAtUtc = string.Empty;
		LastSnapshotPath = string.Empty;
		RevitVersion = string.Empty;
		Summary = new StandardLibrarySnapshotSummary();
	}
}
