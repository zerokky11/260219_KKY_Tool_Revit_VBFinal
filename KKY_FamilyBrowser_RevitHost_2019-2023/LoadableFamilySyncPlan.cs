using System.Collections.Generic;

public class LoadableFamilySyncPlan
{
	public string GeneratedAtUtc { get; set; }

	public string ProjectDocumentTitle { get; set; }

	public string ProjectDocumentPath { get; set; }

	public string StandardDisplayName { get; set; }

	public string ComparisonPath { get; set; }

	public LoadableFamilySyncPlanSummary Summary { get; set; }

	public List<LoadableFamilySyncPlanItem> Items { get; set; }

	public LoadableFamilySyncPlan()
	{
		GeneratedAtUtc = string.Empty;
		ProjectDocumentTitle = string.Empty;
		ProjectDocumentPath = string.Empty;
		StandardDisplayName = string.Empty;
		ComparisonPath = string.Empty;
		Summary = new LoadableFamilySyncPlanSummary();
		Items = new List<LoadableFamilySyncPlanItem>();
	}
}
