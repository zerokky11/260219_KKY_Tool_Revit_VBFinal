public class FamilyThumbnailAutoConfirmedDialogRecord
{
	public string ConfirmedAtUtc { get; set; }

	public string FamilyName { get; set; }

	public string CategoryName { get; set; }

	public string Reason { get; set; }

	public string DialogText { get; set; }

	public string OverrideResult { get; set; }

	public FamilyThumbnailAutoConfirmedDialogRecord()
	{
		ConfirmedAtUtc = string.Empty;
		FamilyName = string.Empty;
		CategoryName = string.Empty;
		Reason = string.Empty;
		DialogText = string.Empty;
		OverrideResult = string.Empty;
	}
}
