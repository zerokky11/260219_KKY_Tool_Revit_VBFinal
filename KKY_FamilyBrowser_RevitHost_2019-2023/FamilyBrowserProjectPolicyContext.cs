public sealed class FamilyBrowserProjectPolicyContext
{
	public string ProjectTitle { get; set; }

	public string ModelPath { get; set; }

	public string CentralPath { get; set; }

	public string StandardTarget { get; set; }

	public bool IsWorkshared { get; set; }

	public FamilyBrowserProjectPolicyContext()
	{
		ProjectTitle = string.Empty;
		ModelPath = string.Empty;
		CentralPath = string.Empty;
		StandardTarget = string.Empty;
		IsWorkshared = false;
	}
}
