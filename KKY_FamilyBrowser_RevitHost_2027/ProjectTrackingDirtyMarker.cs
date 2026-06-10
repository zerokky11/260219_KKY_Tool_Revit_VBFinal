using System.Collections.Generic;

public class ProjectTrackingDirtyMarker
{
	public string DetectedAtUtc { get; set; }

	public string User { get; set; }

	public string DocumentTitle { get; set; }

	public string DocumentPath { get; set; }

	public string State { get; set; }

	public string RequiredAction { get; set; }

	public string Reason { get; set; }

	public List<ProjectTrackingDirtyItem> Items { get; set; }

	public ProjectTrackingDirtyMarker()
	{
		DetectedAtUtc = string.Empty;
		User = string.Empty;
		DocumentTitle = string.Empty;
		DocumentPath = string.Empty;
		State = string.Empty;
		RequiredAction = string.Empty;
		Reason = string.Empty;
		Items = new List<ProjectTrackingDirtyItem>();
	}
}
