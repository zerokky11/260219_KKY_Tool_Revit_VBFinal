using System.Collections.Generic;

public class FamilyThumbnailBatchUpdateItem
{
	public string FamilyName { get; set; }

	public string CategoryName { get; set; }

	public bool Success { get; set; }

	public bool Skipped { get; set; }

	public string ImagePath { get; set; }

	public string Message { get; set; }

	public List<string> AutoConfirmedDialogs { get; set; }

	public FamilyThumbnailBatchUpdateItem()
	{
		FamilyName = string.Empty;
		CategoryName = string.Empty;
		ImagePath = string.Empty;
		Message = string.Empty;
		AutoConfirmedDialogs = new List<string>();
	}
}
