public class SystemSyncExecutionStep
{
	public int SequenceNo { get; set; }

	public string Action { get; set; }

	public string Status { get; set; }

	public string TargetKind { get; set; }

	public string TargetName { get; set; }

	public string Notes { get; set; }

	public SystemSyncExecutionStep()
	{
		Action = string.Empty;
		Status = string.Empty;
		TargetKind = string.Empty;
		TargetName = string.Empty;
		Notes = string.Empty;
	}
}
