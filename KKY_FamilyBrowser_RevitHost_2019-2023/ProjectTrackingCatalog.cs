using System.Collections.Generic;

public class ProjectTrackingCatalog
{
	public string GeneratedAtUtc { get; set; }

	public string SourceId { get; set; }

	public string SourceDisplayName { get; set; }

	public string ApprovedStandardStamp { get; set; }

	public string ProjectDocumentTitle { get; set; }

	public string ProjectDocumentPath { get; set; }

	public List<TrackedLoadableFamilyState> LoadableFamilies { get; set; }

	public List<TrackedSystemTypeState> SystemTypes { get; set; }

	public ProjectTrackingCatalog()
	{
		GeneratedAtUtc = string.Empty;
		SourceId = string.Empty;
		SourceDisplayName = string.Empty;
		ApprovedStandardStamp = string.Empty;
		ProjectDocumentTitle = string.Empty;
		ProjectDocumentPath = string.Empty;
		LoadableFamilies = new List<TrackedLoadableFamilyState>();
		SystemTypes = new List<TrackedSystemTypeState>();
	}
}
