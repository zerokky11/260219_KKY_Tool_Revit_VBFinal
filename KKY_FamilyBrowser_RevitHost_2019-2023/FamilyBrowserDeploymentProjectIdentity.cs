public sealed class FamilyBrowserDeploymentProjectIdentity
{
	public string ProjectTitle { get; set; }

	public string ModelPath { get; set; }

	public string CentralPath { get; set; }

	public FamilyBrowserDeploymentProjectIdentity()
	{
		ProjectTitle = string.Empty;
		ModelPath = string.Empty;
		CentralPath = string.Empty;
	}
}
