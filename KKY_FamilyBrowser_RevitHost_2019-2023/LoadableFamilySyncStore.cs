using System;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;

public sealed class LoadableFamilySyncStore
{
	private LoadableFamilySyncStore()
	{
	}

	public static string Save(string workspaceRoot, LoadableFamilySyncExecutionReport report)
	{
		FamilyBrowserStandardPolicyStore.RequireManagedDataRootForWrite(workspaceRoot, FamilyBrowserLanguageService.Text("Save loadable family apply report", "로더블 패밀리 적용 리포트 저장"));
		string outputDir = Path.Combine(ProjectSnapshotStore.GetProjectHistoryFolder(workspaceRoot, report?.ProjectDocumentPath ?? string.Empty, report?.ProjectDocumentTitle ?? "Untitled"), "LoadableSync");
		Directory.CreateDirectory(outputDir);
		string safeProjectName = MakeSafeFileName(report.ProjectDocumentTitle ?? "Untitled");
		string safeStandardName = MakeSafeFileName(report.StandardDisplayName ?? "StandardLibrary");
		string fileNameStem = "loadable-family-sync-" + safeProjectName + "-" + safeStandardName;
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
