public sealed class FamilyBrowserDeploymentBootstrapResult
{
	public bool Applied { get; set; }

	public bool UsedCache { get; set; }

	public string Source { get; set; }

	public string BootstrapUrl { get; set; }

	public string CachePath { get; set; }

	public string ManagedPolicyPath { get; set; }

	public string ManagedPolicyPathIssue { get; set; }

	public string RequestStorePath { get; set; }

	public string Message { get; set; }

	public string CheckedAtUtc { get; set; }

	public int RefreshMinutes { get; set; }

	public FamilyBrowserDeploymentBootstrapResult()
	{
		Applied = false;
		UsedCache = false;
		Source = string.Empty;
		BootstrapUrl = string.Empty;
		CachePath = string.Empty;
		ManagedPolicyPath = string.Empty;
		ManagedPolicyPathIssue = string.Empty;
		RequestStorePath = string.Empty;
		Message = string.Empty;
		CheckedAtUtc = string.Empty;
		RefreshMinutes = 30;
	}
}
