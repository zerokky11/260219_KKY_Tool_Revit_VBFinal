public class ProjectStandardReference
{
	public string SourceId { get; set; }

	public string DisplayName { get; set; }

	public string SourceKind { get; set; }

	public string ResolvedPath { get; set; }

	public string SnapshotPath { get; set; }

	public string RevitVersion { get; set; }

	public ProjectStandardReference()
	{
		SourceId = string.Empty;
		DisplayName = string.Empty;
		SourceKind = string.Empty;
		ResolvedPath = string.Empty;
		SnapshotPath = string.Empty;
		RevitVersion = string.Empty;
	}
}
