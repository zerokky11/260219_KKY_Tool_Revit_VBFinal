using System.Collections.Generic;

public class ProjectLoadableFamilySnapshotItem
{
	public string FamilyName { get; set; }

	public string CategoryName { get; set; }

	public string CategoryId { get; set; }

	public string CategoryGroup { get; set; }

	public int TypeCount { get; set; }

	public int InstanceCount { get; set; }

	public List<string> TypeNames { get; set; }

	public List<StandardFamilyParameterSnapshotItem> Parameters { get; set; }

	public string ContentFingerprint { get; set; }

	public string ContentSignatureDebugPath { get; set; }

	public string ContentFingerprintFailureReason { get; set; }

	public string UniqueId { get; set; }

	public bool IsShared { get; set; }

	public ProjectLoadableFamilySnapshotItem()
	{
		FamilyName = string.Empty;
		CategoryName = string.Empty;
		CategoryId = string.Empty;
		CategoryGroup = string.Empty;
		TypeNames = new List<string>();
		Parameters = new List<StandardFamilyParameterSnapshotItem>();
		ContentFingerprint = string.Empty;
		ContentSignatureDebugPath = string.Empty;
		ContentFingerprintFailureReason = string.Empty;
		UniqueId = string.Empty;
	}
}
