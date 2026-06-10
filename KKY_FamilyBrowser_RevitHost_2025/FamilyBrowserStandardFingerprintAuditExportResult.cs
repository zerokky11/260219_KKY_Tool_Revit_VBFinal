public sealed class FamilyBrowserStandardFingerprintAuditExportResult
{
	public string OutputPath { get; set; }

	public int RowCount { get; set; }

	public int MissingFingerprintCount { get; set; }

	public int RecoveredFingerprintCount { get; set; }

	public int MissingFromSnapshotCount { get; set; }

	public int SystemTypeRowCount { get; set; }

	public bool WasCanceled { get; set; }

	public string ErrorMessage { get; set; }

	public string SheetName { get; set; }

	public FamilyBrowserStandardFingerprintAuditExportResult()
	{
		OutputPath = string.Empty;
		ErrorMessage = string.Empty;
		SheetName = "FingerprintAudit";
	}
}
