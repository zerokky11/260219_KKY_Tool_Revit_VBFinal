public class LoadableFamilySyncExecutionSummary
{
	public int LoadedCount { get; set; }

	public int ReloadedCount { get; set; }

	public int TrackingRefreshedCount { get; set; }

	public int BlockedCount { get; set; }

	public int SkippedCount { get; set; }

	public int FailedCount { get; set; }
}
