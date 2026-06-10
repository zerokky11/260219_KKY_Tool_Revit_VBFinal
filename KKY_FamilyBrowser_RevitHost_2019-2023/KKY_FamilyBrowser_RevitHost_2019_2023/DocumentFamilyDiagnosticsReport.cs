using System.Collections.Generic;

namespace KKY_FamilyBrowser_RevitHost_2019_2023;

internal class DocumentFamilyDiagnosticsReport
{
	public string GeneratedAtUtc { get; set; }

	public string DocumentTitle { get; set; }

	public string DocumentPath { get; set; }

	public string RevitVersion { get; set; }

	public DocumentFamilyDiagnosticsSummary Summary { get; set; }

	public List<DocumentFamilyDiagnosticsItem> Families { get; set; }

	public DocumentFamilyDiagnosticsReport()
	{
		GeneratedAtUtc = string.Empty;
		DocumentTitle = string.Empty;
		DocumentPath = string.Empty;
		RevitVersion = string.Empty;
		Summary = new DocumentFamilyDiagnosticsSummary();
		Families = new List<DocumentFamilyDiagnosticsItem>();
	}
}
