using System.Collections.Generic;

public class StandardLoadableFamilySnapshotItem
{
	public string FamilyName { get; set; }

	public string CategoryName { get; set; }

	public string CategoryId { get; set; }

	public string CategoryGroup { get; set; }

	public string MetadataMode { get; set; }

	public int TypeCount { get; set; }

	public List<string> TypeNames { get; set; }

	public List<StandardFamilyParameterSnapshotItem> Parameters { get; set; }

	public string ContentFingerprint { get; set; }

	public string ContentSignatureDebugPath { get; set; }

	public string ContentFingerprintFailureReason { get; set; }

	public string UniqueId { get; set; }

	public bool IsShared { get; set; }

	public bool IsNestedLoadableChild { get; set; }

	public bool StandalonePlacementUsageCaptured { get; set; }

	public int StandaloneInstanceCount { get; set; }

	public List<StandardNestedLoadableFamilySnapshotItem> NestedLoadableFamilies { get; set; }

	public string BrowserDetailKey { get; set; }

	public StandardLoadableFamilySnapshotItem()
	{
		FamilyName = string.Empty;
		CategoryName = string.Empty;
		CategoryId = string.Empty;
		CategoryGroup = string.Empty;
		MetadataMode = string.Empty;
		TypeNames = new List<string>();
		Parameters = new List<StandardFamilyParameterSnapshotItem>();
		ContentFingerprint = string.Empty;
		ContentSignatureDebugPath = string.Empty;
		ContentFingerprintFailureReason = string.Empty;
		UniqueId = string.Empty;
		StandalonePlacementUsageCaptured = false;
		StandaloneInstanceCount = 0;
		NestedLoadableFamilies = new List<StandardNestedLoadableFamilySnapshotItem>();
		BrowserDetailKey = string.Empty;
	}
}
