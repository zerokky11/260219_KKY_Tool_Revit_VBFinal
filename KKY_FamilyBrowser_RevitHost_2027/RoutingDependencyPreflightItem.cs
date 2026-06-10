using System.Collections.Generic;

public class RoutingDependencyPreflightItem
{
	public string SystemFamilyKind { get; set; }

	public string SystemTypeName { get; set; }

	public string DependencyRole { get; set; }

	public string SourceLibraryFamilyId { get; set; }

	public string SourceFamilyName { get; set; }

	public string SourceTypeName { get; set; }

	public string SourceFamilyFingerprint { get; set; }

	public string SourceTypeFingerprint { get; set; }

	public string TargetLibraryFamilyId { get; set; }

	public string TargetFamilyName { get; set; }

	public string TargetTypeName { get; set; }

	public string TargetFamilyFingerprint { get; set; }

	public string TargetTypeFingerprint { get; set; }

	public string Action { get; set; }

	public string Reason { get; set; }

	public List<string> RelatedTypeNames { get; set; }

	public RoutingDependencyPreflightItem()
	{
		SystemFamilyKind = string.Empty;
		SystemTypeName = string.Empty;
		DependencyRole = string.Empty;
		SourceLibraryFamilyId = string.Empty;
		SourceFamilyName = string.Empty;
		SourceTypeName = string.Empty;
		SourceFamilyFingerprint = string.Empty;
		SourceTypeFingerprint = string.Empty;
		TargetLibraryFamilyId = string.Empty;
		TargetFamilyName = string.Empty;
		TargetTypeName = string.Empty;
		TargetFamilyFingerprint = string.Empty;
		TargetTypeFingerprint = string.Empty;
		Action = string.Empty;
		Reason = string.Empty;
		RelatedTypeNames = new List<string>();
	}
}
