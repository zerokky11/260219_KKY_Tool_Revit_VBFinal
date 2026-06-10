public class FamilyThumbnailCacheMetadata
{
	public int SchemaVersion { get; set; }

	public string SourceId { get; set; }

	public string SourceFileLastWriteUtc { get; set; }

	public long SourceFileLength { get; set; }

	public string SnapshotMode { get; set; }

	public string SnapshotCapturedAtUtc { get; set; }

	public string CategoryName { get; set; }

	public string FamilyName { get; set; }

	public string FamilyCacheStamp { get; set; }

	public string ImageGeneratedAtUtc { get; set; }

	public FamilyThumbnailCacheMetadata()
	{
		SchemaVersion = 1;
		SourceId = string.Empty;
		SourceFileLastWriteUtc = string.Empty;
		SnapshotMode = string.Empty;
		SnapshotCapturedAtUtc = string.Empty;
		CategoryName = string.Empty;
		FamilyName = string.Empty;
		FamilyCacheStamp = string.Empty;
		ImageGeneratedAtUtc = string.Empty;
	}
}
