using System.Collections.Generic;

public sealed class FamilyBrowserBridgeResponse
{
	public string RequestId { get; set; }

	public bool Success { get; set; }

	public string Message { get; set; }

	public string RevitVersion { get; set; }

	public string ActiveDocumentTitle { get; set; }

	public List<FamilyBrowserBridgeDocumentInfo> Documents { get; set; }

	public FamilyBrowserBridgeResponse()
	{
		RequestId = string.Empty;
		Success = false;
		Message = string.Empty;
		RevitVersion = string.Empty;
		ActiveDocumentTitle = string.Empty;
		Documents = new List<FamilyBrowserBridgeDocumentInfo>();
	}
}
