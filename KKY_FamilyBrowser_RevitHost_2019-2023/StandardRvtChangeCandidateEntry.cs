public class StandardRvtChangeCandidateEntry
{
	public string EntryId { get; set; }

	public string RecordedAtUtc { get; set; }

	public string UserName { get; set; }

	public string DocumentTitle { get; set; }

	public string DocumentPath { get; set; }

	public string SourceId { get; set; }

	public string SlotKey { get; set; }

	public string DisciplineKey { get; set; }

	public string DisciplineLabel { get; set; }

	public string CandidateKind { get; set; }

	public string CategoryName { get; set; }

	public string CategoryId { get; set; }

	public string FamilyName { get; set; }

	public string TypeName { get; set; }

	public string SystemFamilyKind { get; set; }

	public string ChangeKind { get; set; }

	public string Reason { get; set; }

	public string Details { get; set; }

	public StandardRvtChangeCandidateEntry()
	{
		EntryId = string.Empty;
		RecordedAtUtc = string.Empty;
		UserName = string.Empty;
		DocumentTitle = string.Empty;
		DocumentPath = string.Empty;
		SourceId = string.Empty;
		SlotKey = string.Empty;
		DisciplineKey = string.Empty;
		DisciplineLabel = string.Empty;
		CandidateKind = string.Empty;
		CategoryName = string.Empty;
		CategoryId = string.Empty;
		FamilyName = string.Empty;
		TypeName = string.Empty;
		SystemFamilyKind = string.Empty;
		ChangeKind = string.Empty;
		Reason = string.Empty;
		Details = string.Empty;
	}
}
