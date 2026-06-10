public class TrackedSystemTypeState
{
	public string IdentityKey { get; set; }

	public string TypeName { get; set; }

	public string CategoryName { get; set; }

	public string TypeClassName { get; set; }

	public string ApprovedFingerprint { get; set; }

	public string ApprovedStandardStamp { get; set; }

	public string ApprovedAtUtc { get; set; }

	public string ApprovedSemanticFingerprint { get; set; }

	public TrackedSystemTypeState()
	{
		IdentityKey = string.Empty;
		TypeName = string.Empty;
		CategoryName = string.Empty;
		TypeClassName = string.Empty;
		ApprovedFingerprint = string.Empty;
		ApprovedStandardStamp = string.Empty;
		ApprovedAtUtc = string.Empty;
		ApprovedSemanticFingerprint = string.Empty;
	}
}
