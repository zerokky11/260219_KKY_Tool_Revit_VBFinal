public class LoadableFingerprintDifferenceDetailItem
{
	public string Area { get; set; }

	public string DifferenceKind { get; set; }

	public string StandardValue { get; set; }

	public string ProjectValue { get; set; }

	public string Details { get; set; }

	public LoadableFingerprintDifferenceDetailItem()
	{
		Area = string.Empty;
		DifferenceKind = string.Empty;
		StandardValue = string.Empty;
		ProjectValue = string.Empty;
		Details = string.Empty;
	}
}
