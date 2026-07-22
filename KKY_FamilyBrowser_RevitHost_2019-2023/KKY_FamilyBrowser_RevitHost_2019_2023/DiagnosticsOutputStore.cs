using System;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;

namespace KKY_FamilyBrowser_RevitHost_2019_2023;

internal sealed class DiagnosticsOutputStore
{
	private DiagnosticsOutputStore()
	{
	}

	public static string Save(DocumentFamilyDiagnosticsReport report)
	{
		string outputDir = Path.Combine(ResolveWorkspaceRoot(), "Output", "RevitHost");
		Directory.CreateDirectory(outputDir);
		string safeDocumentName = MakeSafeFileName(report.DocumentTitle ?? "Untitled");
		return FamilyBrowserUniqueJsonReportStore.Save(outputDir, "family-browser-diagnostics-" + safeDocumentName, report);
	}

	private static string ResolveWorkspaceRoot()
	{
		return HostWorkspacePathResolver.ResolveRoot();
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
