public class SystemTypePreflightDiagnostic
{
	public string Stage { get; set; }

	public string PlannedAction { get; set; }

	public string NormalizedAction { get; set; }

	public string SystemTypeName { get; set; }

	public string SystemFamilyKind { get; set; }

	public string Reason { get; set; }

	public string Details { get; set; }

	public SystemTypePreflightDiagnostic()
	{
		Stage = string.Empty;
		PlannedAction = string.Empty;
		NormalizedAction = string.Empty;
		SystemTypeName = string.Empty;
		SystemFamilyKind = string.Empty;
		Reason = string.Empty;
		Details = string.Empty;
	}
}
