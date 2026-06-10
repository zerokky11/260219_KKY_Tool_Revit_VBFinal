using System.Collections.Generic;

public sealed class FamilyBrowserDeploymentBootstrapSecurityRefreshResult
{
	public bool Success { get; set; }

	public bool Changed { get; set; }

	public string Source { get; set; }

	public string BootstrapUrl { get; set; }

	public string Message { get; set; }

	public string CheckedAtUtc { get; set; }

	public int RefreshMinutes { get; set; }

	public List<string> AdminProfileKeywords { get; set; }

	public FamilyBrowserDeploymentBootstrapSecurityRefreshResult()
	{
		Success = false;
		Changed = false;
		Source = string.Empty;
		BootstrapUrl = string.Empty;
		Message = string.Empty;
		CheckedAtUtc = string.Empty;
		RefreshMinutes = 30;
		AdminProfileKeywords = new List<string>();
	}
}
