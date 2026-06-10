using System;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;

namespace KKY_FamilyBrowser_RevitHost_2027;

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
		string fileName = "family-browser-diagnostics-" + safeDocumentName + "-" + DateTime.Now.ToString("yyyyMMdd-HHmmss", CultureInfo.InvariantCulture) + ".json";
		string text = Path.Combine(outputDir, fileName);
		string json = PlainJsonReportWriter.Serialize(report);
		File.WriteAllText(text, json);
		return text;
	}

	private static string ResolveWorkspaceRoot()
	{
		return HostWorkspacePathResolver.ResolveRoot();
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
