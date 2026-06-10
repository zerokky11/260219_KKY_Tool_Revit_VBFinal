using System.Collections.Generic;

public class SystemSyncExecutionPlan
{
	public string GeneratedAtUtc { get; set; }

	public List<SystemSyncExecutionItem> Items { get; set; }

	public SystemSyncExecutionPlan()
	{
		GeneratedAtUtc = string.Empty;
		Items = new List<SystemSyncExecutionItem>();
	}
}
