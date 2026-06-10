using System;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;

public sealed class SystemTypeApplyStore
{
	private SystemTypeApplyStore()
	{
	}

	public static string Save(string workspaceRoot, SystemTypeApplyExecutionReport report)
	{
		FamilyBrowserStandardPolicyStore.RequireManagedDataRootForWrite(workspaceRoot, FamilyBrowserLanguageService.Text("Save system type apply report", "시스템 타입 적용 리포트 저장"));
		string outputDir = Path.Combine(ProjectSnapshotStore.GetProjectHistoryFolder(workspaceRoot, report?.ProjectDocumentPath ?? string.Empty, report?.ProjectDocumentTitle ?? "Untitled"), "SystemTypeSync");
		Directory.CreateDirectory(outputDir);
		string safeProjectName = MakeSafeFileName(report.ProjectDocumentTitle ?? "Untitled");
		string safeStandardName = MakeSafeFileName(report.StandardDisplayName ?? "StandardLibrary");
		string fileName = "system-type-sync-" + safeProjectName + "-" + safeStandardName + "-" + DateTime.Now.ToString("yyyyMMdd-HHmmss", CultureInfo.InvariantCulture) + ".json";
		string text = Path.Combine(outputDir, fileName);
		File.WriteAllText(text, PlainJsonReportWriter.Serialize(report));
		return text;
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
