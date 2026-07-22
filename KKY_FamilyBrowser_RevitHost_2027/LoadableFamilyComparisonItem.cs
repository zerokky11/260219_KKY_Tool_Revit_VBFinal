using System.Collections.Generic;

public class LoadableFamilyComparisonItem
{
	public string IdentityKey { get; set; }

	public string FamilyName { get; set; }

	public string CategoryName { get; set; }

	public string Status { get; set; }

	public string StandardFingerprint { get; set; }

	public string ProjectFingerprint { get; set; }

	public string StandardContentFingerprint { get; set; }

	public string ProjectContentFingerprint { get; set; }

	public string StandardContentSignatureDebugPath { get; set; }

	public string ProjectContentSignatureDebugPath { get; set; }

	public string StandardContentFingerprintFailureReason { get; set; }

	public string ProjectContentFingerprintFailureReason { get; set; }

	public List<string> FingerprintDifferenceSummary { get; set; }

	public List<LoadableFingerprintDifferenceDetailItem> FingerprintDifferenceDetails { get; set; }

	public int StandardTypeCount { get; set; }

	public int ProjectTypeCount { get; set; }

	public int ProjectInstanceCount { get; set; }

	public List<string> MissingTypeNames { get; set; }

	public List<string> ExtraTypeNames { get; set; }

	public List<string> ProjectTypeNames { get; set; }

	public List<StandardFamilyParameterSnapshotItem> ProjectParameters { get; set; }

	public List<StandardNestedLoadableFamilySnapshotItem> ProjectNestedLoadableFamilies { get; set; }

	public string Notes { get; set; }

	public bool IsNestedLoadableChild { get; set; }

	public bool IsNestedLoadableDifference { get; set; }

	public List<string> NestedParentFamilyNames { get; set; }

	public List<string> NestedDifferenceFamilyNames { get; set; }

	public List<StandardNestedLoadableFamilySnapshotItem> NestedLoadableFamilies { get; set; }

	public LoadableFamilyComparisonItem()
	{
		IdentityKey = string.Empty;
		FamilyName = string.Empty;
		CategoryName = string.Empty;
		Status = string.Empty;
		StandardFingerprint = string.Empty;
		ProjectFingerprint = string.Empty;
		StandardContentFingerprint = string.Empty;
		ProjectContentFingerprint = string.Empty;
		StandardContentSignatureDebugPath = string.Empty;
		ProjectContentSignatureDebugPath = string.Empty;
		StandardContentFingerprintFailureReason = string.Empty;
		ProjectContentFingerprintFailureReason = string.Empty;
		FingerprintDifferenceSummary = new List<string>();
		FingerprintDifferenceDetails = new List<LoadableFingerprintDifferenceDetailItem>();
		MissingTypeNames = new List<string>();
		ExtraTypeNames = new List<string>();
		ProjectTypeNames = new List<string>();
		ProjectParameters = new List<StandardFamilyParameterSnapshotItem>();
		ProjectNestedLoadableFamilies = new List<StandardNestedLoadableFamilySnapshotItem>();
		Notes = string.Empty;
		NestedParentFamilyNames = new List<string>();
		NestedDifferenceFamilyNames = new List<string>();
		NestedLoadableFamilies = new List<StandardNestedLoadableFamilySnapshotItem>();
	}
}
