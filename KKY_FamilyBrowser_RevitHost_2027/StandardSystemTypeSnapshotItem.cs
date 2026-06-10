using System.Collections.Generic;

public class StandardSystemTypeSnapshotItem
{
	public string TypeName { get; set; }

	public string CategoryName { get; set; }

	public string CategoryId { get; set; }

	public string TypeClassName { get; set; }

	public string UniqueId { get; set; }

	public bool SupportsRoutingDependencies { get; set; }

	public string SemanticFingerprint { get; set; }

	public string ClassificationCode { get; set; }

	public string SegmentName { get; set; }

	public string MaterialName { get; set; }

	public string Shape { get; set; }

	public string RoutingPreferenceSignature { get; set; }

	public string CompoundStructureSignature { get; set; }

	public List<StandardSystemTypeLayerSnapshotItem> Layers { get; set; }

	public StandardSystemTypeSnapshotItem()
	{
		TypeName = string.Empty;
		CategoryName = string.Empty;
		CategoryId = string.Empty;
		TypeClassName = string.Empty;
		UniqueId = string.Empty;
		SemanticFingerprint = string.Empty;
		ClassificationCode = string.Empty;
		SegmentName = string.Empty;
		MaterialName = string.Empty;
		Shape = string.Empty;
		RoutingPreferenceSignature = string.Empty;
		CompoundStructureSignature = string.Empty;
		Layers = new List<StandardSystemTypeLayerSnapshotItem>();
	}
}
