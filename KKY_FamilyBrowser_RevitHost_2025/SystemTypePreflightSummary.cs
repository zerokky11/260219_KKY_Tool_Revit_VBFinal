public class SystemTypePreflightSummary
{
	public int NoChangeCount { get; set; }

	public int ReadyCount { get; set; }

	public int ApprovalRequiredCount { get; set; }

	public int BlockedCount { get; set; }

	public int LoadableFoundationBlockedCount { get; set; }

	public int MissingDependencyFamilyCount { get; set; }

	public int DependencyReloadCount { get; set; }

	public int DependencyManualReviewCount { get; set; }
}
