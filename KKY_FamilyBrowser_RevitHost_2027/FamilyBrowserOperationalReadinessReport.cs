using System.Collections.Generic;

public sealed class FamilyBrowserOperationalReadinessReport
{
	public string GeneratedAtUtc { get; set; }

	public string PolicyPath { get; set; }

	public string StandardTarget { get; set; }

	public int BlockingCount { get; set; }

	public int WarningCount { get; set; }

	public int ReadyCount { get; set; }

	public List<FamilyBrowserOperationalReadinessItem> Items { get; set; }

	public bool IsReadyForModeler => BlockingCount == 0;

	public FamilyBrowserOperationalReadinessReport()
	{
		GeneratedAtUtc = string.Empty;
		PolicyPath = string.Empty;
		StandardTarget = string.Empty;
		BlockingCount = 0;
		WarningCount = 0;
		ReadyCount = 0;
		Items = new List<FamilyBrowserOperationalReadinessItem>();
	}
}
