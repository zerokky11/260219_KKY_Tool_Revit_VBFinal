using System.Collections.Generic;

public class ProjectStandardComparisonReport
{
	public string IdentityMode { get; set; }

	public string TrackingState { get; set; }

	public string GeneratedAtUtc { get; set; }

	public ProjectStandardReference Standard { get; set; }

	public ProjectReference Project { get; set; }

	public ProjectStandardComparisonSummary Summary { get; set; }

	public List<ProjectLoadableSignatureFailureItem> ProjectLoadableSignatureFailures { get; set; }

	public List<LoadableFamilyComparisonItem> LoadableFamilies { get; set; }

	public List<SystemTypeComparisonItem> SystemTypes { get; set; }

	public ProjectStandardComparisonReport()
	{
		IdentityMode = string.Empty;
		TrackingState = string.Empty;
		GeneratedAtUtc = string.Empty;
		Standard = new ProjectStandardReference();
		Project = new ProjectReference();
		Summary = new ProjectStandardComparisonSummary();
		ProjectLoadableSignatureFailures = new List<ProjectLoadableSignatureFailureItem>();
		LoadableFamilies = new List<LoadableFamilyComparisonItem>();
		SystemTypes = new List<SystemTypeComparisonItem>();
	}
}
