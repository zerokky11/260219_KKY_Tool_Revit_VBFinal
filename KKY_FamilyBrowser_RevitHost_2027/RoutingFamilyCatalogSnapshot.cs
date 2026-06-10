using System.Collections.Generic;

public class RoutingFamilyCatalogSnapshot
{
	public string SourceId { get; set; }

	public string DocumentTitle { get; set; }

	public string DocumentPath { get; set; }

	public string CapturedAtUtc { get; set; }

	public List<RoutingFamilyCatalogEntry> Families { get; set; }

	public RoutingFamilyCatalogSnapshot()
	{
		SourceId = string.Empty;
		DocumentTitle = string.Empty;
		DocumentPath = string.Empty;
		CapturedAtUtc = string.Empty;
		Families = new List<RoutingFamilyCatalogEntry>();
	}
}
