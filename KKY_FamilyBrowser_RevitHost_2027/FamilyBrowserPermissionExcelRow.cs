using System;
using System.Collections.Generic;

internal sealed class FamilyBrowserPermissionExcelRow
{
	public int RowNumber { get; set; }

	public string Enabled { get; set; }

	public string ApplyFolder { get; set; }

	public string RvtFileName { get; set; }

	public string ProjectKey { get; set; }

	public string ProjectName { get; set; }

	public string Discipline { get; set; }

	public string MatchMode { get; set; }

	public string MatchValue { get; set; }

	public string CentralPath { get; set; }

	public string CentralPathContains { get; set; }

	public string ModelPath { get; set; }

	public string ModelPathContains { get; set; }

	public string ProjectTitleContains { get; set; }

	public string UserOrGroup { get; set; }

	public string Role { get; set; }

	public Dictionary<string, string> Permissions { get; set; }

	public FamilyBrowserPermissionExcelRow()
	{
		RowNumber = 0;
		Enabled = string.Empty;
		ApplyFolder = string.Empty;
		RvtFileName = string.Empty;
		ProjectKey = string.Empty;
		ProjectName = string.Empty;
		Discipline = string.Empty;
		MatchMode = string.Empty;
		MatchValue = string.Empty;
		CentralPath = string.Empty;
		CentralPathContains = string.Empty;
		ModelPath = string.Empty;
		ModelPathContains = string.Empty;
		ProjectTitleContains = string.Empty;
		UserOrGroup = string.Empty;
		Role = string.Empty;
		Permissions = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
	}
}
