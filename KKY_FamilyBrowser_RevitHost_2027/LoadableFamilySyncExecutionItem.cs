public class LoadableFamilySyncExecutionItem
{
	public string IdentityKey { get; set; }

	public string FamilyName { get; set; }

	public string CategoryName { get; set; }

	public string ComparisonStatus { get; set; }

	public string PlannedAction { get; set; }

	public string ExecutionMode { get; set; }

	public string Outcome { get; set; }

	public string Details { get; set; }

	public LoadableFamilySyncExecutionItem()
	{
		IdentityKey = string.Empty;
		FamilyName = string.Empty;
		CategoryName = string.Empty;
		ComparisonStatus = string.Empty;
		PlannedAction = string.Empty;
		ExecutionMode = string.Empty;
		Outcome = string.Empty;
		Details = string.Empty;
	}
}
