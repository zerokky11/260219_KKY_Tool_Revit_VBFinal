public class FamilyBrowserOperationLogEntry
{
	public string EntryId { get; set; }

	public string RecordedAtUtc { get; set; }

	public string UserName { get; set; }

	public string OperationKind { get; set; }

	public string DocumentTitle { get; set; }

	public string DocumentPath { get; set; }

	public string SourceId { get; set; }

	public string StandardSourceKey { get; set; }

	public string StandardDisplayName { get; set; }

	public string CandidateKind { get; set; }

	public string CategoryName { get; set; }

	public string FamilyName { get; set; }

	public string TypeName { get; set; }

	public string SystemFamilyKind { get; set; }

	public string PlannedAction { get; set; }

	public string Outcome { get; set; }

	public string Details { get; set; }

	public string CommitState { get; set; }

	public string CommitKind { get; set; }

	public string CommittedAtUtc { get; set; }

	public FamilyBrowserOperationLogEntry()
	{
		EntryId = string.Empty;
		RecordedAtUtc = string.Empty;
		UserName = string.Empty;
		OperationKind = string.Empty;
		DocumentTitle = string.Empty;
		DocumentPath = string.Empty;
		SourceId = string.Empty;
		StandardSourceKey = string.Empty;
		StandardDisplayName = string.Empty;
		CandidateKind = string.Empty;
		CategoryName = string.Empty;
		FamilyName = string.Empty;
		TypeName = string.Empty;
		SystemFamilyKind = string.Empty;
		PlannedAction = string.Empty;
		Outcome = string.Empty;
		Details = string.Empty;
		CommitState = string.Empty;
		CommitKind = string.Empty;
		CommittedAtUtc = string.Empty;
	}
}
