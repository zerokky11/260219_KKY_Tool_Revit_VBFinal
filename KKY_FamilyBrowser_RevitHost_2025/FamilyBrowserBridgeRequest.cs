public sealed class FamilyBrowserBridgeRequest
{
	public string RequestId { get; set; }

	public string Command { get; set; }

	public string TargetDocumentId { get; set; }

	public string PayloadPath { get; set; }

	public string CreatedBy { get; set; }

	public FamilyBrowserBridgeRequest()
	{
		RequestId = string.Empty;
		Command = string.Empty;
		TargetDocumentId = string.Empty;
		PayloadPath = string.Empty;
		CreatedBy = string.Empty;
	}
}
