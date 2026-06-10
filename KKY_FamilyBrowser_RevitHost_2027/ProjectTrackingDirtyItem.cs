public class ProjectTrackingDirtyItem
{
	public string Action { get; set; }

	public string Kind { get; set; }

	public string Name { get; set; }

	public string CategoryName { get; set; }

	public string ElementIdText { get; set; }

	public string State { get; set; }

	public string RecoveryStatus { get; set; }

	public string RequiredAction { get; set; }

	public ProjectTrackingDirtyItem()
	{
		Action = string.Empty;
		Kind = string.Empty;
		Name = string.Empty;
		CategoryName = string.Empty;
		ElementIdText = string.Empty;
		State = string.Empty;
		RecoveryStatus = string.Empty;
		RequiredAction = string.Empty;
	}
}
