public sealed class FamilyBrowserOperationalReadinessItem
{
	public string Area { get; set; }

	public string Status { get; set; }

	public string Message { get; set; }

	public string Action { get; set; }

	public FamilyBrowserOperationalReadinessItem()
	{
		Area = string.Empty;
		Status = "Ready";
		Message = string.Empty;
		Action = string.Empty;
	}
}
