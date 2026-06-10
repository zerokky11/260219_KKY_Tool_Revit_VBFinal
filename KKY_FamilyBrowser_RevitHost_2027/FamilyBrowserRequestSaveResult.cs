public class FamilyBrowserRequestSaveResult
{
	public string RequestPath { get; set; }

	public string MailDraftPath { get; set; }

	public string AttachmentFolder { get; set; }

	public int AttachmentCount { get; set; }

	public string StoreMode { get; set; }

	public string StoreLocation { get; set; }

	public string ConnectorNote { get; set; }

	public FamilyBrowserRequestSaveResult()
	{
		RequestPath = string.Empty;
		MailDraftPath = string.Empty;
		AttachmentFolder = string.Empty;
		AttachmentCount = 0;
		StoreMode = string.Empty;
		StoreLocation = string.Empty;
		ConnectorNote = string.Empty;
	}
}
