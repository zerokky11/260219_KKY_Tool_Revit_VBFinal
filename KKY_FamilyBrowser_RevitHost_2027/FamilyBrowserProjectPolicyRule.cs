using System.Collections.Generic;

public sealed class FamilyBrowserProjectPolicyRule
{
	public string RuleName { get; set; }

	public bool Enabled { get; set; }

	public string MatchMode { get; set; }

	public string MatchValue { get; set; }

	public string PermissionPreset { get; set; }

	public List<string> CustomAdminUsers { get; set; }

	public List<string> CustomRequestApproverUsers { get; set; }

	public List<string> CustomReadOnlyUsers { get; set; }

	public string AllowUnlistedUsersAsModelers { get; set; }

	public string AllowModelersToLoadFamilies { get; set; }

	public string AllowModelersToApplySystemTypes { get; set; }

	public string AllowModelersToSubmitRequests { get; set; }

	public string LastUpdatedUtc { get; set; }

	public string LastUpdatedBy { get; set; }

	public FamilyBrowserProjectPolicyRule()
	{
		RuleName = string.Empty;
		Enabled = true;
		MatchMode = "CentralPathContains";
		MatchValue = string.Empty;
		PermissionPreset = "Inherit";
		CustomAdminUsers = new List<string>();
		CustomRequestApproverUsers = new List<string>();
		CustomReadOnlyUsers = new List<string>();
		AllowUnlistedUsersAsModelers = "Inherit";
		AllowModelersToLoadFamilies = "Inherit";
		AllowModelersToApplySystemTypes = "Inherit";
		AllowModelersToSubmitRequests = "Inherit";
		LastUpdatedUtc = string.Empty;
		LastUpdatedBy = string.Empty;
	}
}
