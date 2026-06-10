using System.Collections.Generic;

public class SystemTypeSyncPlanItem
{
	public string SystemFamilyKind { get; set; }

	public string CategoryName { get; set; }

	public string SourceTypeName { get; set; }

	public string DestinationTypeName { get; set; }

	public string Action { get; set; }

	public string Reason { get; set; }

	public string SourceFingerprint { get; set; }

	public string DestinationFingerprint { get; set; }

	public List<string> DiffSummary { get; set; }

	public List<string> RelatedDuplicateNames { get; set; }

	public SystemTypeSyncPlanItem()
	{
		SystemFamilyKind = string.Empty;
		CategoryName = string.Empty;
		SourceTypeName = string.Empty;
		DestinationTypeName = string.Empty;
		Action = string.Empty;
		Reason = string.Empty;
		SourceFingerprint = string.Empty;
		DestinationFingerprint = string.Empty;
		DiffSummary = new List<string>();
		RelatedDuplicateNames = new List<string>();
	}
}
