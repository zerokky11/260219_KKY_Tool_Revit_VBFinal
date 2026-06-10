using System.Collections.Generic;

public class RoutingFamilyCatalogEntry
{
	public string LibraryFamilyId { get; set; }

	public string FamilyName { get; set; }

	public string FamilyFingerprint { get; set; }

	public List<RoutingFamilyTypeSnapshot> Types { get; set; }

	public RoutingFamilyCatalogEntry()
	{
		LibraryFamilyId = string.Empty;
		FamilyName = string.Empty;
		FamilyFingerprint = string.Empty;
		Types = new List<RoutingFamilyTypeSnapshot>();
	}
}
