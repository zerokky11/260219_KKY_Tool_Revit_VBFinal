using System.Collections.Generic;

public class LoadableFamilySyncExecutionReport
{
	public string GeneratedAtUtc { get; set; }

	public string ProjectDocumentTitle { get; set; }

	public string ProjectDocumentPath { get; set; }

	public string StandardDisplayName { get; set; }

	public string ComparisonPath { get; set; }

	public string PostComparisonPath { get; set; }

	public string TrackingPath { get; set; }

	public LoadableFamilySyncExecutionSummary Summary { get; set; }

	public List<LoadableFamilySyncExecutionItem> Items { get; set; }

	public LoadableFamilySyncExecutionReport()
	{
		GeneratedAtUtc = string.Empty;
		ProjectDocumentTitle = string.Empty;
		ProjectDocumentPath = string.Empty;
		StandardDisplayName = string.Empty;
		ComparisonPath = string.Empty;
		PostComparisonPath = string.Empty;
		TrackingPath = string.Empty;
		Summary = new LoadableFamilySyncExecutionSummary();
		Items = new List<LoadableFamilySyncExecutionItem>();
	}
}
