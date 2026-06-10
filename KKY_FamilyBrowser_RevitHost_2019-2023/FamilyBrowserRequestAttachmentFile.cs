public class FamilyBrowserRequestAttachmentFile
{
	public string DisplayName { get; set; }

	public string OriginalPath { get; set; }

	public string StoredPath { get; set; }

	public string RelativePath { get; set; }

	public long SizeBytes { get; set; }

	public string AttachedAtUtc { get; set; }

	public string AttachedBy { get; set; }

	public FamilyBrowserRequestAttachmentFile()
	{
		DisplayName = string.Empty;
		OriginalPath = string.Empty;
		StoredPath = string.Empty;
		RelativePath = string.Empty;
		SizeBytes = 0L;
		AttachedAtUtc = string.Empty;
		AttachedBy = string.Empty;
	}
}
