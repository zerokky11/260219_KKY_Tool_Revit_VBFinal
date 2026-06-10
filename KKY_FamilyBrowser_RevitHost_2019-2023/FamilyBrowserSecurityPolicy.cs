using System.Collections.Generic;

public sealed class FamilyBrowserSecurityPolicy
{
	public List<string> AdminUsers { get; set; }

	public List<string> AdminProfileKeywords { get; set; }

	public List<string> RequestApproverUsers { get; set; }

	public List<string> ReadOnlyUsers { get; set; }

	public bool AllowUnlistedUsersAsModelers { get; set; }

	public bool AllowModelersToLoadFamilies { get; set; }

	public bool AllowModelersToApplySystemTypes { get; set; }

	public bool AllowModelersToSubmitRequests { get; set; }

	public string LastUpdatedUtc { get; set; }

	public string LastUpdatedBy { get; set; }

	public FamilyBrowserSecurityPolicy()
	{
		AdminUsers = new List<string>();
		AdminProfileKeywords = new List<string>();
		RequestApproverUsers = new List<string>();
		ReadOnlyUsers = new List<string>();
		AllowUnlistedUsersAsModelers = true;
		AllowModelersToLoadFamilies = true;
		AllowModelersToApplySystemTypes = true;
		AllowModelersToSubmitRequests = true;
		LastUpdatedUtc = string.Empty;
		LastUpdatedBy = string.Empty;
	}

	public static FamilyBrowserSecurityPolicy CreateDefault()
	{
		return new FamilyBrowserSecurityPolicy();
	}
}
