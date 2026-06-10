public class ProjectLoadableSignatureFailureItem
{
	public string FamilyName { get; set; }

	public string CategoryName { get; set; }

	public string CategoryId { get; set; }

	public string CategoryGroup { get; set; }

	public int TypeCount { get; set; }

	public int InstanceCount { get; set; }

	public string FailureKind { get; set; }

	public string Reason { get; set; }

	public string ContentFingerprint { get; set; }

	public string ContentSignatureDebugPath { get; set; }

	public string UniqueId { get; set; }

	public bool IsShared { get; set; }

	public ProjectLoadableSignatureFailureItem()
	{
		FamilyName = string.Empty;
		CategoryName = string.Empty;
		CategoryId = string.Empty;
		CategoryGroup = string.Empty;
		FailureKind = string.Empty;
		Reason = string.Empty;
		ContentFingerprint = string.Empty;
		ContentSignatureDebugPath = string.Empty;
		UniqueId = string.Empty;
	}
}
