using System;
using System.Collections.Generic;

public class SystemTypeSemanticSnapshot
{
	public string SystemFamilyKind { get; set; }

	public string CategoryName { get; set; }

	public string TypeName { get; set; }

	public string ClassificationCode { get; set; }

	public string SegmentName { get; set; }

	public string MaterialName { get; set; }

	public string Shape { get; set; }

	public string RoutingPreferenceSignature { get; set; }

	public string CompoundStructureSignature { get; set; }

	public Dictionary<string, string> Parameters { get; set; }

	public List<RoutingDependencySnapshot> RoutingDependencies { get; set; }

	public bool DetailedComponentsCaptured { get; set; }

	public string DetailedComponentSignature { get; set; }

	public List<SystemTypeDetailedComponentSnapshotItem> DetailedComponents { get; set; }

	public string DetailSummary { get; set; }

	public SystemTypeSemanticSnapshot()
	{
		SystemFamilyKind = string.Empty;
		CategoryName = string.Empty;
		TypeName = string.Empty;
		ClassificationCode = string.Empty;
		SegmentName = string.Empty;
		MaterialName = string.Empty;
		Shape = string.Empty;
		RoutingPreferenceSignature = string.Empty;
		CompoundStructureSignature = string.Empty;
		Parameters = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
		RoutingDependencies = new List<RoutingDependencySnapshot>();
		DetailedComponentsCaptured = false;
		DetailedComponentSignature = string.Empty;
		DetailedComponents = new List<SystemTypeDetailedComponentSnapshotItem>();
		DetailSummary = string.Empty;
	}
}
