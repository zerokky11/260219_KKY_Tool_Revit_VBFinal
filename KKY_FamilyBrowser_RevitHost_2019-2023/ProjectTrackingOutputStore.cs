using System;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;

public sealed class ProjectTrackingOutputStore
{
	private ProjectTrackingOutputStore()
	{
	}

	public static string Save(string workspaceRoot, ProjectTrackingCatalog catalog)
	{
		FamilyBrowserStandardPolicyStore.RequireManagedDataRootForWrite(workspaceRoot, FamilyBrowserLanguageService.Text("Save project tracking report", "프로젝트 추적 리포트 저장"));
		string outputDir = Path.Combine(ProjectSnapshotStore.GetProjectHistoryFolder(workspaceRoot, catalog?.ProjectDocumentPath ?? string.Empty, catalog?.ProjectDocumentTitle ?? "Untitled"), "Tracking");
		Directory.CreateDirectory(outputDir);
		string safeProjectName = MakeSafeFileName(catalog.ProjectDocumentTitle ?? "Untitled");
		string fileNameStem = "project-tracking-" + safeProjectName;
		return FamilyBrowserUniqueJsonReportStore.Save(outputDir, fileNameStem, catalog);
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
