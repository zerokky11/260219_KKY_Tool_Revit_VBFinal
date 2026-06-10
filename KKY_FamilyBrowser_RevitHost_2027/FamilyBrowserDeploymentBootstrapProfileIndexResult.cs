using System.Collections.Generic;

public sealed class FamilyBrowserDeploymentBootstrapProfileIndexResult
{
	public bool Success { get; set; }

	public string IndexUrl { get; set; }

	public string Source { get; set; }

	public string Message { get; set; }

	public string DefaultProfileId { get; set; }

	public List<FamilyBrowserDeploymentBootstrapProfile> Profiles { get; set; }

	public List<FamilyBrowserDeploymentBootstrapProjectRule> ProjectRules { get; set; }

	public FamilyBrowserDeploymentBootstrapProfileIndexResult()
	{
		Success = false;
		IndexUrl = string.Empty;
		Source = string.Empty;
		Message = string.Empty;
		DefaultProfileId = string.Empty;
		Profiles = new List<FamilyBrowserDeploymentBootstrapProfile>();
		ProjectRules = new List<FamilyBrowserDeploymentBootstrapProjectRule>();
	}
}
