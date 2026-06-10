public class FamilyThumbnailPreviewResult
{
	public bool Success { get; set; }

	public string ImagePath { get; set; }

	public string Message { get; set; }

	public FamilyThumbnailPreviewResult()
	{
		ImagePath = string.Empty;
		Message = string.Empty;
	}
}
