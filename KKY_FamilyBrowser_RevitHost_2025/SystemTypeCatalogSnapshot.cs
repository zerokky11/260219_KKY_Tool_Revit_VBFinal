using System.Collections.Generic;

public class SystemTypeCatalogSnapshot
{
	public string SourceId { get; set; }

	public string DocumentTitle { get; set; }

	public string DocumentPath { get; set; }

	public string CapturedAtUtc { get; set; }

	public string RevitVersion { get; set; }

	public List<SystemTypeSemanticSnapshot> Types { get; set; }

	public SystemTypeCatalogSnapshot()
	{
		SourceId = string.Empty;
		DocumentTitle = string.Empty;
		DocumentPath = string.Empty;
		CapturedAtUtc = string.Empty;
		RevitVersion = string.Empty;
		Types = new List<SystemTypeSemanticSnapshot>();
	}
}
