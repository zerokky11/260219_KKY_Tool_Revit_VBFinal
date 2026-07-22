using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Windows.Forms;

internal static class FamilyBrowserResultExcelExportUi
{
	public static bool SaveRows(
		IWin32Window owner,
		FamilyBrowserHtmlDialogHost host,
		bool isKorean,
		string suggestedFileName,
		string sheetName,
		IList<string> headers,
		IList<List<string>> rows)
	{
		using (SaveFileDialog dialog = new SaveFileDialog())
		{
			dialog.Title = isKorean ? "결과 Excel 내보내기" : "Export result to Excel";
			dialog.Filter = "Excel Workbook (*.xlsx)|*.xlsx";
			dialog.DefaultExt = "xlsx";
			dialog.AddExtension = true;
			dialog.OverwritePrompt = true;
			dialog.CheckPathExists = true;
			dialog.RestoreDirectory = true;
			dialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
			dialog.FileName = EnsureXlsxFileName(suggestedFileName);
			if (dialog.ShowDialog(owner) != DialogResult.OK)
			{
				return false;
			}

			try
			{
				FamilyBrowserStandardListExcelExportResult result = FamilyBrowserStandardListExcelExportService.SaveRows(dialog.FileName, sheetName, headers, rows);
				if (host != null)
				{
					host.SetStatusMessage((isKorean ? "Excel 내보내기 완료: " : "Excel exported: ") + result.OutputPath);
				}
				return true;
			}
			catch (Exception ex)
			{
				if (host != null)
				{
					host.SetStatusMessage((isKorean ? "Excel 내보내기 실패: " : "Excel export failed: ") + ex.Message);
				}
				return false;
			}
		}
	}

	public static string TimestampedFileName(string prefix)
	{
		return SafeFileName(prefix) + "-" + DateTime.Now.ToString("yyyyMMdd-HHmmss", CultureInfo.InvariantCulture) + ".xlsx";
	}

	private static string EnsureXlsxFileName(string fileName)
	{
		string value = SafeFileName(string.IsNullOrWhiteSpace(fileName) ? "KKY-FamilyBrowser-Result" : Path.GetFileNameWithoutExtension(fileName));
		return value + ".xlsx";
	}

	private static string SafeFileName(string value)
	{
		string safe = string.IsNullOrWhiteSpace(value) ? "KKY-FamilyBrowser-Result" : value.Trim();
		foreach (char invalid in Path.GetInvalidFileNameChars())
		{
			safe = safe.Replace(invalid, '_');
		}
		return safe;
	}
}
