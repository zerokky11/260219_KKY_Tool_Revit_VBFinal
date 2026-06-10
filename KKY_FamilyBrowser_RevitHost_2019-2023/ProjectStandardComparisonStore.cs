using System;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;

public sealed class ProjectStandardComparisonStore
{
	private ProjectStandardComparisonStore()
	{
	}

	public static string Save(string workspaceRoot, ProjectStandardComparisonReport report)
	{
		FamilyBrowserStandardPolicyStore.RequireManagedDataRootForWrite(workspaceRoot, FamilyBrowserLanguageService.Text("Save project comparison report", "프로젝트 비교 리포트 저장"));
		string outputDir = Path.Combine(ProjectSnapshotStore.GetProjectHistoryFolder(workspaceRoot, report?.Project?.DocumentPath ?? string.Empty, report?.Project?.DocumentTitle ?? string.Empty), "Comparisons");
		Directory.CreateDirectory(outputDir);
		string safeProjectName = MakeSafeFileName(report.Project.DocumentTitle ?? "Untitled");
		string safeStandardName = MakeSafeFileName(report.Standard.DisplayName ?? "StandardLibrary");
		string fileName = "project-vs-standard-" + safeProjectName + "-" + safeStandardName + "-" + DateTime.Now.ToString("yyyyMMdd-HHmmss", CultureInfo.InvariantCulture) + ".json";
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
