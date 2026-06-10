public class ProjectReference
{
	public string DocumentTitle { get; set; }

	public string DocumentPath { get; set; }

	public string RevitVersion { get; set; }

	public string SnapshotPath { get; set; }

	public string ThumbnailSourceId { get; set; }

	public string ThumbnailFolder { get; set; }

	public ProjectReference()
	{
		DocumentTitle = string.Empty;
		DocumentPath = string.Empty;
		RevitVersion = string.Empty;
		SnapshotPath = string.Empty;
		ThumbnailSourceId = string.Empty;
		ThumbnailFolder = string.Empty;
	}
}
