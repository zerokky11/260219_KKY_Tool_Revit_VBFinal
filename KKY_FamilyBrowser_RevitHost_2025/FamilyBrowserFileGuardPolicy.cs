using System.Collections.Generic;

public sealed class FamilyBrowserFileGuardPolicy
{
	public bool Enabled { get; set; }

	public string RootFolder { get; set; }

	public List<FamilyBrowserFileGuardTarget> Targets { get; set; }

	public string LastUpdatedUtc { get; set; }

	public string LastUpdatedBy { get; set; }

	public FamilyBrowserFileGuardPolicy()
	{
		Enabled = false;
		RootFolder = string.Empty;
		Targets = new List<FamilyBrowserFileGuardTarget>();
		LastUpdatedUtc = string.Empty;
		LastUpdatedBy = string.Empty;
	}

	public static FamilyBrowserFileGuardPolicy CreateDefault()
	{
		return new FamilyBrowserFileGuardPolicy();
	}
}
