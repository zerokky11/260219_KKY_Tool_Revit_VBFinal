using System.Collections.Generic;

public class RoutingDependencyPreflightPlan
{
	public string GeneratedAtUtc { get; set; }

	public List<RoutingDependencyPreflightItem> Items { get; set; }

	public RoutingDependencyPreflightPlan()
	{
		GeneratedAtUtc = string.Empty;
		Items = new List<RoutingDependencyPreflightItem>();
	}
}
