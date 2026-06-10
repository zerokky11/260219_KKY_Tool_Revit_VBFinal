using System.Collections.Generic;

public class FamilyBrowserOperationLog
{
	public int SchemaVersion { get; set; }

	public string LogDate { get; set; }

	public string UpdatedAtUtc { get; set; }

	public List<FamilyBrowserOperationLogEntry> Entries { get; set; }

	public FamilyBrowserOperationLog()
	{
		SchemaVersion = 1;
		LogDate = string.Empty;
		UpdatedAtUtc = string.Empty;
		Entries = new List<FamilyBrowserOperationLogEntry>();
	}
}
