public sealed class LoadableSignatureDebugRecord
{
	public string CategoryName { get; set; }

	public string FamilyName { get; set; }

	public string Fingerprint { get; set; }

	public string ErrorMessage { get; set; }

	public string Mode { get; set; }

	public string Path { get; set; }

	public long LastWriteUtcTicks { get; set; }

	public LoadableSignatureDebugRecord()
	{
		CategoryName = string.Empty;
		FamilyName = string.Empty;
		Fingerprint = string.Empty;
		ErrorMessage = string.Empty;
		Mode = string.Empty;
		Path = string.Empty;
	}
}
