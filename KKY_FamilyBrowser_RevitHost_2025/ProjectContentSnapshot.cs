using System.Collections.Generic;

public class ProjectContentSnapshot
{
	public string DocumentTitle { get; set; }

	public string DocumentPath { get; set; }

	public string CapturedAtUtc { get; set; }

	public string RevitVersion { get; set; }

	public ProjectContentSnapshotSummary Summary { get; set; }

	public List<ProjectLoadableFamilySnapshotItem> LoadableFamilies { get; set; }

	public List<ProjectLoadableSignatureFailureItem> LoadableSignatureFailures { get; set; }

	public List<ProjectSystemTypeSnapshotItem> SystemTypes { get; set; }

	public ProjectContentSnapshot()
	{
		DocumentTitle = string.Empty;
		DocumentPath = string.Empty;
		CapturedAtUtc = string.Empty;
		RevitVersion = string.Empty;
		Summary = new ProjectContentSnapshotSummary();
		LoadableFamilies = new List<ProjectLoadableFamilySnapshotItem>();
		LoadableSignatureFailures = new List<ProjectLoadableSignatureFailureItem>();
		SystemTypes = new List<ProjectSystemTypeSnapshotItem>();
	}
}
