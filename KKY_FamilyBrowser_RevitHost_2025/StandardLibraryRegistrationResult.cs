using System.Collections.Generic;

public class StandardLibraryRegistrationResult
{
	public StandardLibraryRegistrationRecord Registration { get; set; }

	public StandardLibrarySnapshot Snapshot { get; set; }

	public string RegistrationPath { get; set; }

	public string SnapshotPath { get; set; }

	public string DiagnosticReportPath { get; set; }

	public List<FamilyThumbnailAutoConfirmedDialogRecord> AutoHandledDialogs { get; set; }

	public StandardLibraryRegistrationResult()
	{
		Registration = new StandardLibraryRegistrationRecord();
		Snapshot = new StandardLibrarySnapshot();
		RegistrationPath = string.Empty;
		SnapshotPath = string.Empty;
		DiagnosticReportPath = string.Empty;
		AutoHandledDialogs = new List<FamilyThumbnailAutoConfirmedDialogRecord>();
	}
}
