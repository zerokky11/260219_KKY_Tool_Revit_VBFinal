using System.Collections.Generic;

public class SystemTypeSyncPlan
{
	public string GeneratedAtUtc { get; set; }

	public List<SystemTypeSyncPlanItem> Items { get; set; }

	public SystemTypeSyncPlan()
	{
		GeneratedAtUtc = string.Empty;
		Items = new List<SystemTypeSyncPlanItem>();
	}
}
