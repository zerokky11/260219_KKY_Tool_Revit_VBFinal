public class LoadableFamilyContentSignatureResult
{
	public string Fingerprint { get; set; }

	public string Signature { get; set; }

	public string DebugMetadata { get; set; }

	public string Mode { get; set; }

	public string ErrorMessage { get; set; }

	public LoadableFamilyContentSignatureResult()
	{
		Fingerprint = string.Empty;
		Signature = string.Empty;
		DebugMetadata = string.Empty;
		Mode = string.Empty;
		ErrorMessage = string.Empty;
	}
}
