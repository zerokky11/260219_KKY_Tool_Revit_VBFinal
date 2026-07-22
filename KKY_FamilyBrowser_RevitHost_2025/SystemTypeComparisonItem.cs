using System.Collections.Generic;

public class SystemTypeComparisonItem
{
	public string IdentityKey { get; set; }

	public string TypeName { get; set; }

	public string CategoryName { get; set; }

	public string TypeClassName { get; set; }

	public string Status { get; set; }

	public string StandardFingerprint { get; set; }

	public string ProjectFingerprint { get; set; }

	public List<string> DifferenceSummary { get; set; }

	public bool SupportsRoutingDependencies { get; set; }

	public string StandardRoutingPreferenceSignature { get; set; }

	public string ProjectRoutingPreferenceSignature { get; set; }

	public string Notes { get; set; }

	public string DetailSummary { get; set; }

	public List<StandardSystemTypeLayerSnapshotItem> Layers { get; set; }

	public SystemTypeComparisonItem()
	{
		IdentityKey = string.Empty;
		TypeName = string.Empty;
		CategoryName = string.Empty;
		TypeClassName = string.Empty;
		Status = string.Empty;
		StandardFingerprint = string.Empty;
		ProjectFingerprint = string.Empty;
		DifferenceSummary = new List<string>();
		StandardRoutingPreferenceSignature = string.Empty;
		ProjectRoutingPreferenceSignature = string.Empty;
		Notes = string.Empty;
		DetailSummary = string.Empty;
		Layers = new List<StandardSystemTypeLayerSnapshotItem>();
	}
}
