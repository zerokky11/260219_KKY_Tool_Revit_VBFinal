public class RoutingDependencySnapshot
{
	public string DependencyRole { get; set; }

	public string LibraryFamilyId { get; set; }

	public string FamilyName { get; set; }

	public string TypeName { get; set; }

	public string FamilyFingerprint { get; set; }

	public string TypeFingerprint { get; set; }

	public RoutingDependencySnapshot()
	{
		DependencyRole = string.Empty;
		LibraryFamilyId = string.Empty;
		FamilyName = string.Empty;
		TypeName = string.Empty;
		FamilyFingerprint = string.Empty;
		TypeFingerprint = string.Empty;
	}
}
