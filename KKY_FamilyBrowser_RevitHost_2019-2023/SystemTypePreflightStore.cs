using System;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;

public sealed class SystemTypePreflightStore
{
	private SystemTypePreflightStore()
	{
	}

	public static string Save(string workspaceRoot, SystemTypePreflightReport report)
	{
		FamilyBrowserStandardPolicyStore.RequireManagedDataRootForWrite(workspaceRoot, FamilyBrowserLanguageService.Text("Save system type preflight report", "시스템 타입 사전검토 리포트 저장"));
		string outputDir = Path.Combine(ProjectSnapshotStore.GetProjectHistoryFolder(workspaceRoot, report?.ProjectDocumentPath ?? string.Empty, report?.ProjectDocumentTitle ?? "Untitled"), "SystemTypePreflight");
		Directory.CreateDirectory(outputDir);
		string safeProjectName = MakeSafeFileName(report.ProjectDocumentTitle ?? "Untitled");
		string safeStandardName = MakeSafeFileName(report.StandardDisplayName ?? "StandardLibrary");
		string fileNameStem = "system-type-preflight-" + safeProjectName + "-" + safeStandardName;
		return FamilyBrowserUniqueJsonReportStore.Save(outputDir, fileNameStem, report);
	}

	private static string MakeSafeFileName(string value)
	{
		char[] invalidFileNameChars = Path.GetInvalidFileNameChars();
		string normalized = new string(value.Select([SpecialName] (char ch) => (!invalidFileNameChars.Contains(ch)) ? ch : '_').ToArray()).Trim();
		if (normalized.Length == 0)
		{
			return "Untitled";
		}
		return normalized;
	}
}
