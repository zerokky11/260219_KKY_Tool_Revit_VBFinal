public class StandardLibrarySnapshotCacheHit
{
	public StandardLibrarySnapshot Snapshot { get; set; }

	public string SnapshotPath { get; set; }

	public StandardLibrarySnapshotCacheHit()
	{
		Snapshot = new StandardLibrarySnapshot();
		SnapshotPath = string.Empty;
	}
}
