public class ProjectScanCacheRecord
{
	public int SchemaVersion { get; set; }

	public string ProjectKey { get; set; }

	public string ProjectTitle { get; set; }

	public string ProjectDocumentPath { get; set; }

	public string ProjectCentralPath { get; set; }

	public string ProjectIdentityPath { get; set; }

	public string ProjectDocumentRevisionToken { get; set; }

	public bool CapturedFromModifiedDocument { get; set; }

	public string ProjectFileLastWriteUtc { get; set; }

	public long ProjectFileLength { get; set; }

	public string RevitVersion { get; set; }

	public string StandardSourceId { get; set; }

	public string StandardSnapshotPath { get; set; }

	public string StandardSnapshotAtUtc { get; set; }

	public string StandardSnapshotMode { get; set; }

	public string StandardSourceFileLastWriteUtc { get; set; }

	public long StandardSourceFileLength { get; set; }

	public string ProjectSnapshotPath { get; set; }

	public string ComparisonReportPath { get; set; }

	public string ThumbnailSourceId { get; set; }

	public string ThumbnailFolder { get; set; }

	public string CapturedAtUtc { get; set; }

	public string SavedAtUtc { get; set; }

	public ProjectScanCacheRecord()
	{
		SchemaVersion = 4;
		ProjectKey = string.Empty;
		ProjectTitle = string.Empty;
		ProjectDocumentPath = string.Empty;
		ProjectCentralPath = string.Empty;
		ProjectIdentityPath = string.Empty;
		ProjectDocumentRevisionToken = string.Empty;
		ProjectFileLastWriteUtc = string.Empty;
		RevitVersion = string.Empty;
		StandardSourceId = string.Empty;
		StandardSnapshotPath = string.Empty;
		StandardSnapshotAtUtc = string.Empty;
		StandardSnapshotMode = string.Empty;
		StandardSourceFileLastWriteUtc = string.Empty;
		ProjectSnapshotPath = string.Empty;
		ComparisonReportPath = string.Empty;
		ThumbnailSourceId = string.Empty;
		ThumbnailFolder = string.Empty;
		CapturedAtUtc = string.Empty;
		SavedAtUtc = string.Empty;
	}
}
