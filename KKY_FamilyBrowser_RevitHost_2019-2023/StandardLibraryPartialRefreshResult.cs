using System.Collections.Generic;

public class StandardLibraryPartialRefreshResult
{
	public StandardLibraryRegistrationRecord Registration { get; set; }

	public StandardLibrarySnapshot Snapshot { get; set; }

	public string RegistrationPath { get; set; }

	public string SnapshotPath { get; set; }

	public int RequestedCount { get; set; }

	public int UpdatedCount { get; set; }

	public List<string> MissingFamilyNames { get; set; }

	public string DiagnosticReportPath { get; set; }

	public List<FamilyThumbnailAutoConfirmedDialogRecord> AutoHandledDialogs { get; set; }

	public StandardLibraryPartialRefreshResult()
	{
		Registration = new StandardLibraryRegistrationRecord();
		Snapshot = new StandardLibrarySnapshot();
		RegistrationPath = string.Empty;
		SnapshotPath = string.Empty;
		MissingFamilyNames = new List<string>();
		DiagnosticReportPath = string.Empty;
		AutoHandledDialogs = new List<FamilyThumbnailAutoConfirmedDialogRecord>();
	}
}
