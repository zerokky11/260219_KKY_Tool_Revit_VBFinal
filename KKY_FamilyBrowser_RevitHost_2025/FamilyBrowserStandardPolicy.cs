using System.Collections.Generic;

public sealed class FamilyBrowserStandardPolicy
{
	public string Mode { get; set; }

	public string ActiveDiscipline { get; set; }

	public FamilyBrowserStandardLibrarySlot IntegratedLibrary { get; set; }

	public List<FamilyBrowserStandardLibrarySlot> DisciplineLibraries { get; set; }

	public FamilyBrowserRequestStoreSettings RequestStore { get; set; }

	public FamilyBrowserPermissionExcelSettings PermissionExcel { get; set; }

	public FamilyBrowserSecurityPolicy Security { get; set; }

	public List<FamilyBrowserProjectPolicyRule> ProjectPolicyRules { get; set; }

	public FamilyBrowserFileGuardPolicy FileGuard { get; set; }

	public string LastUpdatedUtc { get; set; }

	public string LastUpdatedBy { get; set; }

	public FamilyBrowserStandardPolicy()
	{
		Mode = "DisciplineSeparated";
		ActiveDiscipline = "Mechanical";
		IntegratedLibrary = FamilyBrowserStandardLibrarySlot.CreateIntegrated();
		DisciplineLibraries = FamilyBrowserStandardLibrarySlot.CreateDefaultDisciplines();
		RequestStore = FamilyBrowserRequestStoreSettings.CreateDefault();
		PermissionExcel = FamilyBrowserPermissionExcelSettings.CreateDefault();
		Security = FamilyBrowserSecurityPolicy.CreateDefault();
		ProjectPolicyRules = new List<FamilyBrowserProjectPolicyRule>();
		FileGuard = FamilyBrowserFileGuardPolicy.CreateDefault();
		LastUpdatedUtc = string.Empty;
		LastUpdatedBy = string.Empty;
	}
}
