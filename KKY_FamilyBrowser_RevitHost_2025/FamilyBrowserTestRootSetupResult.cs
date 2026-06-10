public sealed class FamilyBrowserTestRootSetupResult
{
	public string RootFolder { get; set; }

	public string ConfigFolder { get; set; }

	public string RequestFolder { get; set; }

	public string StandardsFolder { get; set; }

	public string RegistryFolder { get; set; }

	public string SnapshotFolder { get; set; }

	public string ThumbnailFolder { get; set; }

	public string SharedPolicyPath { get; set; }

	public string GuidePath { get; set; }

	public bool Writable { get; set; }

	public string WritableDetail { get; set; }

	public bool TeamVisiblePath { get; set; }

	public FamilyBrowserTestRootSetupResult()
	{
		RootFolder = string.Empty;
		ConfigFolder = string.Empty;
		RequestFolder = string.Empty;
		StandardsFolder = string.Empty;
		RegistryFolder = string.Empty;
		SnapshotFolder = string.Empty;
		ThumbnailFolder = string.Empty;
		SharedPolicyPath = string.Empty;
		GuidePath = string.Empty;
		Writable = false;
		WritableDetail = string.Empty;
		TeamVisiblePath = false;
	}
}
