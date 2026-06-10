using System.Collections.Generic;

public class SystemTypeApplyExecutionItem
{
	public string IdentityKey { get; set; }

	public string SystemFamilyKind { get; set; }

	public string CategoryName { get; set; }

	public string SystemTypeName { get; set; }

	public string SyncAction { get; set; }

	public string PreflightStatus { get; set; }

	public string Outcome { get; set; }

	public string AppliedTypeName { get; set; }

	public string BackupTypeName { get; set; }

	public int RetypedElementCount { get; set; }

	public int DeletedObsoleteTypeCount { get; set; }

	public List<string> DependencyActions { get; set; }

	public List<string> Messages { get; set; }

	public string Details { get; set; }

	public SystemTypeApplyExecutionItem()
	{
		IdentityKey = string.Empty;
		SystemFamilyKind = string.Empty;
		CategoryName = string.Empty;
		SystemTypeName = string.Empty;
		SyncAction = string.Empty;
		PreflightStatus = string.Empty;
		Outcome = string.Empty;
		AppliedTypeName = string.Empty;
		BackupTypeName = string.Empty;
		DependencyActions = new List<string>();
		Messages = new List<string>();
		Details = string.Empty;
	}
}
