using System;
using System.Collections.Generic;

public sealed class FamilyBrowserPermissionExcelDiagnostic
{
	public bool Enabled { get; set; }

	public string SourcePath { get; set; }

	public string SheetName { get; set; }

	public bool Exists { get; set; }

	public int RowCount { get; set; }

	public int ActiveRowCount { get; set; }

	public int ProjectMatchedRowCount { get; set; }

	public int UserMatchedRowCount { get; set; }

	public string CurrentUser { get; set; }

	public string ProjectTitle { get; set; }

	public string ModelPath { get; set; }

	public string CentralPath { get; set; }

	public string StandardTarget { get; set; }

	public bool Matched { get; set; }

	public int MatchedRowNumber { get; set; }

	public string MatchedRole { get; set; }

	public string MatchedUser { get; set; }

	public string MatchedProjectName { get; set; }

	public string MatchedDiscipline { get; set; }

	public string MatchedMode { get; set; }

	public string MatchedValue { get; set; }

	public string Message { get; set; }

	public string LastError { get; set; }

	public Dictionary<string, string> PermissionTokens { get; set; }

	public FamilyBrowserPermissionExcelDiagnostic()
	{
		Enabled = false;
		SourcePath = string.Empty;
		SheetName = string.Empty;
		Exists = false;
		RowCount = 0;
		ActiveRowCount = 0;
		ProjectMatchedRowCount = 0;
		UserMatchedRowCount = 0;
		CurrentUser = string.Empty;
		ProjectTitle = string.Empty;
		ModelPath = string.Empty;
		CentralPath = string.Empty;
		StandardTarget = string.Empty;
		Matched = false;
		MatchedRowNumber = 0;
		MatchedRole = string.Empty;
		MatchedUser = string.Empty;
		MatchedProjectName = string.Empty;
		MatchedDiscipline = string.Empty;
		MatchedMode = string.Empty;
		MatchedValue = string.Empty;
		Message = string.Empty;
		LastError = string.Empty;
		PermissionTokens = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
	}
}
