public sealed class FamilyBrowserMachineConfig
{
	public bool UseManagedPolicy { get; set; }

	public string ManagedPolicyPath { get; set; }

	public string DeploymentBootstrapUrl { get; set; }

	public string LastBootstrapCheckUtc { get; set; }

	public string LastBootstrapStatus { get; set; }

	public string LastBootstrapSource { get; set; }

	public string LastUpdatedUtc { get; set; }

	public string LastUpdatedBy { get; set; }

	public FamilyBrowserMachineConfig()
	{
		UseManagedPolicy = false;
		ManagedPolicyPath = string.Empty;
		DeploymentBootstrapUrl = string.Empty;
		LastBootstrapCheckUtc = string.Empty;
		LastBootstrapStatus = string.Empty;
		LastBootstrapSource = string.Empty;
		LastUpdatedUtc = string.Empty;
		LastUpdatedBy = string.Empty;
	}

	public static FamilyBrowserMachineConfig CreateDefault()
	{
		return new FamilyBrowserMachineConfig();
	}
}
