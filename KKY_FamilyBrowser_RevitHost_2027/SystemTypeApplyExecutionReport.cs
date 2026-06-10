using System.Collections.Generic;

public class SystemTypeApplyExecutionReport
{
	public string GeneratedAtUtc { get; set; }

	public string ProjectDocumentTitle { get; set; }

	public string ProjectDocumentPath { get; set; }

	public string StandardDisplayName { get; set; }

	public string PreflightPath { get; set; }

	public string PostPreflightPath { get; set; }

	public string TrackingPath { get; set; }

	public SystemTypeApplyExecutionSummary Summary { get; set; }

	public List<SystemTypeApplyExecutionItem> Items { get; set; }

	public SystemTypeApplyExecutionReport()
	{
		GeneratedAtUtc = string.Empty;
		ProjectDocumentTitle = string.Empty;
		ProjectDocumentPath = string.Empty;
		StandardDisplayName = string.Empty;
		PreflightPath = string.Empty;
		PostPreflightPath = string.Empty;
		TrackingPath = string.Empty;
		Summary = new SystemTypeApplyExecutionSummary();
		Items = new List<SystemTypeApplyExecutionItem>();
	}
}
