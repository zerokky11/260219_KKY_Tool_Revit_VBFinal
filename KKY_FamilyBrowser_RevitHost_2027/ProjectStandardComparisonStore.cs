using System;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;

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
		string fileName = "project-vs-standard-" + safeProjectName + "-" + safeStandardName + "-" + DateTime.UtcNow.ToString("yyyyMMdd-HHmmss-fffffff", CultureInfo.InvariantCulture) + "-" + Guid.NewGuid().ToString("N").Substring(0, 8) + ".json";
		string text = Path.Combine(outputDir, fileName);
		WriteJsonAtomic(text, report);
		return text;
	}

	private static void WriteJsonAtomic(string path, object value)
	{
		string temporaryPath = FamilyBrowserAtomicFileService.CreateSiblingTemporaryPath(path);
		try
		{
			File.WriteAllText(temporaryPath, PlainJsonReportWriter.Serialize(value), new UTF8Encoding(false));
			FamilyBrowserAtomicFileService.Promote(temporaryPath, path);
		}
		finally
		{
			if (File.Exists(temporaryPath))
			{
				File.Delete(temporaryPath);
			}
		}
	}

	private static string MakeSafeFileName(string value)
	{
		char[] invalidFileNameChars = Path.GetInvalidFileNameChars();
		string normalized = new string(value.Select([SpecialName] (char ch) => (!Enumerable.Contains(invalidFileNameChars, ch)) ? ch : '_').ToArray()).Trim();
		if (normalized.Length == 0)
		{
			return "Untitled";
		}
		return normalized;
	}
}
