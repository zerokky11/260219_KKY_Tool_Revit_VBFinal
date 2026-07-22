using System.Collections.Generic;

public class FamilyBrowserRequestRecord
{
	public string RequestId { get; set; }

	public string RequestKind { get; set; }

	public string Status { get; set; }

	public string CreatedAtUtc { get; set; }

	public string CreatedBy { get; set; }

	public string UpdatedAtUtc { get; set; }

	public string LastUpdatedBy { get; set; }

	public long Revision { get; set; }

	public string RevisionToken { get; set; }

	public string ProjectTitle { get; set; }

	public string ProjectPath { get; set; }

	public string CentralPath { get; set; }

	public string StandardTarget { get; set; }

	public string StandardMode { get; set; }

	public string StandardDisplayName { get; set; }

	public string StandardRvtPath { get; set; }

	public string StandardSourceId { get; set; }

	public string ComparisonPath { get; set; }

	public string PreflightPath { get; set; }

	public string TrackingState { get; set; }

	public string ItemName { get; set; }

	public string CategoryName { get; set; }

	public string Discipline { get; set; }

	public string SuggestedAction { get; set; }

	public string Reason { get; set; }

	public string Notes { get; set; }

	public string ProgressNote { get; set; }

	public string SourcePath { get; set; }

	public string AttachmentFolder { get; set; }

	public List<string> Attachments { get; set; }

	public List<FamilyBrowserRequestAttachmentFile> AttachmentFiles { get; set; }

	public List<string> History { get; set; }

	public FamilyBrowserRequestRecord()
	{
		RequestId = string.Empty;
		RequestKind = string.Empty;
		Status = "Draft";
		CreatedAtUtc = string.Empty;
		CreatedBy = string.Empty;
		UpdatedAtUtc = string.Empty;
		LastUpdatedBy = string.Empty;
		Revision = 0L;
		RevisionToken = string.Empty;
		ProjectTitle = string.Empty;
		ProjectPath = string.Empty;
		CentralPath = string.Empty;
		StandardTarget = string.Empty;
		StandardMode = string.Empty;
		StandardDisplayName = string.Empty;
		StandardRvtPath = string.Empty;
		StandardSourceId = string.Empty;
		ComparisonPath = string.Empty;
		PreflightPath = string.Empty;
		TrackingState = string.Empty;
		ItemName = string.Empty;
		CategoryName = string.Empty;
		Discipline = string.Empty;
		SuggestedAction = string.Empty;
		Reason = string.Empty;
		Notes = string.Empty;
		ProgressNote = string.Empty;
		SourcePath = string.Empty;
		AttachmentFolder = string.Empty;
		Attachments = new List<string>();
		AttachmentFiles = new List<FamilyBrowserRequestAttachmentFile>();
		History = new List<string>();
	}
}
