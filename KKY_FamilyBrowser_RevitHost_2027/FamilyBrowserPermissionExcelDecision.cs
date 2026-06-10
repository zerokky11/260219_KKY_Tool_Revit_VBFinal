public sealed class FamilyBrowserPermissionExcelDecision
{
	public bool HasDecision { get; set; }

	public bool Allowed { get; set; }

	public string Role { get; set; }

	public string SourcePath { get; set; }

	public int SourceRow { get; set; }

	public string Message { get; set; }

	public FamilyBrowserPermissionExcelDecision()
	{
		HasDecision = false;
		Allowed = false;
		Role = string.Empty;
		SourcePath = string.Empty;
		SourceRow = 0;
		Message = string.Empty;
	}
}
