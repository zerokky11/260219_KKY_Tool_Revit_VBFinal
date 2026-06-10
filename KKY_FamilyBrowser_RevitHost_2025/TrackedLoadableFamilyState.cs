public class TrackedLoadableFamilyState
{
	public string IdentityKey { get; set; }

	public string FamilyName { get; set; }

	public string CategoryName { get; set; }

	public string ApprovedFingerprint { get; set; }

	public string ApprovedStandardStamp { get; set; }

	public string ApprovedAtUtc { get; set; }

	public TrackedLoadableFamilyState()
	{
		IdentityKey = string.Empty;
		FamilyName = string.Empty;
		CategoryName = string.Empty;
		ApprovedFingerprint = string.Empty;
		ApprovedStandardStamp = string.Empty;
		ApprovedAtUtc = string.Empty;
	}
}
