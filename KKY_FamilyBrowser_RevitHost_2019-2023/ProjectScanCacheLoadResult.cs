public class ProjectScanCacheLoadResult
{
	public bool Success { get; set; }

	public string Reason { get; set; }

	public ProjectScanCacheRecord Record { get; set; }

	public ProjectContentSnapshot Snapshot { get; set; }

	public ProjectStandardComparisonReport Report { get; set; }

	public ProjectScanCacheLoadResult()
	{
		Reason = string.Empty;
	}
}
