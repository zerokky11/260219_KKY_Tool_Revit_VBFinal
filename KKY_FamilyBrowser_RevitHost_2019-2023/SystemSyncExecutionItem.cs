using System.Collections.Generic;

public class SystemSyncExecutionItem
{
	public string SystemFamilyKind { get; set; }

	public string CategoryName { get; set; }

	public string SystemTypeName { get; set; }

	public string SyncAction { get; set; }

	public string ExecutionStatus { get; set; }

	public bool RequiresApproval { get; set; }

	public bool HasManualReview { get; set; }

	public bool RequiresLoadableFoundation { get; set; }

	public string Summary { get; set; }

	public List<string> BlockingReasons { get; set; }

	public List<string> FoundationBlockingReasons { get; set; }

	public List<string> RelatedDuplicateNames { get; set; }

	public List<SystemSyncExecutionStep> Steps { get; set; }

	public SystemSyncExecutionItem()
	{
		SystemFamilyKind = string.Empty;
		CategoryName = string.Empty;
		SystemTypeName = string.Empty;
		SyncAction = string.Empty;
		ExecutionStatus = string.Empty;
		Summary = string.Empty;
		BlockingReasons = new List<string>();
		FoundationBlockingReasons = new List<string>();
		RelatedDuplicateNames = new List<string>();
		Steps = new List<SystemSyncExecutionStep>();
	}
}
