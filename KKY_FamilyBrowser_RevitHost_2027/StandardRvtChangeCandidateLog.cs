using System.Collections.Generic;

public class StandardRvtChangeCandidateLog
{
	public int SchemaVersion { get; set; }

	public string SourceId { get; set; }

	public string StandardRvtPath { get; set; }

	public string UpdatedAtUtc { get; set; }

	public List<StandardRvtChangeCandidateEntry> Entries { get; set; }

	public StandardRvtChangeCandidateLog()
	{
		SchemaVersion = 1;
		SourceId = string.Empty;
		StandardRvtPath = string.Empty;
		UpdatedAtUtc = string.Empty;
		Entries = new List<StandardRvtChangeCandidateEntry>();
	}
}
