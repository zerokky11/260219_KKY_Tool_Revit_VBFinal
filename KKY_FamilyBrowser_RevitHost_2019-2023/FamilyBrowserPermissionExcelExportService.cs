using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Security;
using System.Text;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;

public sealed class FamilyBrowserPermissionExcelExportService
{
	[CompilerGenerated]
	internal sealed class _Closure_0024__3_002D0
	{
		public int _0024VB_0024Local_skipped;

		public _Closure_0024__3_002D0(_Closure_0024__3_002D0 arg0)
		{
			if (arg0 != null)
			{
				_0024VB_0024Local_skipped = arg0._0024VB_0024Local_skipped;
			}
		}

		[SpecialName]
		internal bool _Lambda_0024__0(string path)
		{
			checked
			{
				if (IsLikelyRevitBackup(path))
				{
					_0024VB_0024Local_skipped++;
					return false;
				}
				return true;
			}
		}
	}

	public const string DefaultSheetName = "Policy";

	private FamilyBrowserPermissionExcelExportService()
	{
	}

	public static FamilyBrowserPermissionExcelExportResult SaveTemplate(string outputPath, string defaultDiscipline, string currentUser, bool korean = false)
	{
		List<List<string>> rows = new List<List<string>>
		{
			BuildTemplateRow("X", "\\\\server\\BIM\\Project_A", "Central.rvt", "*", "X", "X", korean ? "샘플 행입니다. 실제 적용 전 사용을 O로 바꾸세요." : "Sample row. Change Enabled to O after editing."),
			BuildTemplateRow("X", "\\\\server\\BIM\\Project_A", "Central.rvt", string.IsNullOrWhiteSpace(currentUser) ? "DOMAIN\\modeler" : currentUser, "O", "X", korean ? "특정 사용자에게 패밀리 로드/편집만 허용하는 예시입니다." : "Example allowing family load/edit for one user.")
		};
		string sheetName = ResolveSheetName(korean);
		WriteWorkbook(outputPath, sheetName, BuildHeaders(korean), rows);
		return new FamilyBrowserPermissionExcelExportResult
		{
			OutputPath = outputPath,
			RowCount = rows.Count,
			SheetName = sheetName
		};
	}

	public static FamilyBrowserPermissionExcelExportResult ExportRvtFolder(string sourceFolder, string outputPath, string defaultDiscipline, bool korean = false)
	{
		_Closure_0024__3_002D0 arg = default(_Closure_0024__3_002D0);
		_Closure_0024__3_002D0 CS_0024_003C_003E8__locals3 = new _Closure_0024__3_002D0(arg);
		if (string.IsNullOrWhiteSpace(sourceFolder) || !Directory.Exists(sourceFolder))
		{
			throw new DirectoryNotFoundException("RVT source folder was not found: " + (sourceFolder ?? string.Empty));
		}
		CS_0024_003C_003E8__locals3._0024VB_0024Local_skipped = 0;
		List<string> rvtFiles = Enumerable.Where(Directory.EnumerateFiles(sourceFolder, "*.rvt", SearchOption.AllDirectories), checked([SpecialName] (string path) =>
		{
			if (IsLikelyRevitBackup(path))
			{
				CS_0024_003C_003E8__locals3._0024VB_0024Local_skipped++;
				return false;
			}
			return true;
		})).OrderBy([SpecialName] (string path) => path, StringComparer.OrdinalIgnoreCase).ToList();
		List<List<string>> rows = new List<List<string>>();
		foreach (string rvtPath in rvtFiles)
		{
			rows.Add(BuildRvtPolicyRow(sourceFolder, rvtPath, korean));
		}
		string sheetName = ResolveSheetName(korean);
		WriteWorkbook(outputPath, sheetName, BuildHeaders(korean), rows);
		return new FamilyBrowserPermissionExcelExportResult
		{
			OutputPath = outputPath,
			SourceFolder = sourceFolder,
			RowCount = rows.Count,
			SkippedBackupCount = CS_0024_003C_003E8__locals3._0024VB_0024Local_skipped,
			SheetName = sheetName
		};
	}

	public static FamilyBrowserPermissionExcelExportResult ExportFileGuardPolicy(FamilyBrowserFileGuardPolicy fileGuard, string outputPath, bool korean = false)
	{
		FamilyBrowserFileGuardPolicy guard = fileGuard ?? FamilyBrowserFileGuardPolicy.CreateDefault();
		List<FamilyBrowserFileGuardTarget> targets = (guard.Targets ?? new List<FamilyBrowserFileGuardTarget>()).Where([SpecialName] (FamilyBrowserFileGuardTarget x) => x != null).OrderBy([SpecialName] (FamilyBrowserFileGuardTarget x) => x.RelativePath ?? string.Empty, StringComparer.OrdinalIgnoreCase).ThenBy([SpecialName] (FamilyBrowserFileGuardTarget x) => x.FileName ?? string.Empty, StringComparer.OrdinalIgnoreCase)
			.ToList();
		List<List<string>> rows = new List<List<string>>();
		foreach (FamilyBrowserFileGuardTarget target in targets)
		{
			rows.Add(BuildFileGuardRow(guard, target));
		}
		string sheetName = (korean ? DecodeUtf8("7YyM7J2867OE6raM7ZWc") : "FileGuards");
		WriteWorkbook(outputPath, sheetName, BuildFileGuardHeaders(korean), rows);
		return new FamilyBrowserPermissionExcelExportResult
		{
			OutputPath = outputPath,
			SourceFolder = (guard.RootFolder ?? string.Empty),
			RowCount = rows.Count,
			SheetName = sheetName
		};
	}

	private static List<string> BuildHeaders(bool korean)
	{
		if (korean)
		{
			return new List<string> { "사용", "적용 폴더", "RVT 파일명", "사용자", "패밀리 로드/편집", "타입 추가/삭제", "비고" };
		}
		return new List<string> { "Enabled", "ApplyFolder", "RvtFileName", "User", "LoadEditFamily", "AddDeleteType", "Notes" };
	}

	private static string ResolveSheetName(bool korean)
	{
		if (!korean)
		{
			return "Policy";
		}
		return "권한";
	}

	private static List<string> BuildFileGuardHeaders(bool korean)
	{
		if (korean)
		{
			return new List<string>
			{
				DecodeUtf8("7KCB7Jqp"),
				DecodeUtf8("7KCB7JqpIO2PtOuNlA=="),
				DecodeUtf8("UlZUIO2MjOydvA=="),
				DecodeUtf8("7IOB64yAIOqyveuhnA=="),
				DecodeUtf8("7KSR7JWZIOqyveuhnA=="),
				"공종",
				DecodeUtf8("7JqU7IaMIOyDneyEscK37IiY7KCVwrfsgq3soJwg7LaU7KCB"),
				DecodeUtf8("7Yyo67CA66asIOuhnOuTnC/tjrjsp5Eg7LCo64uo"),
				DecodeUtf8("7YOA7J6FIOuzgOqyvSDssKjri6g="),
				DecodeUtf8("7ZWY7JyEIOyghOyaqSDtjKjrsIDrpqwg64uo64+FIOuqqOuNuOungSDquIjsp4A="),
				DecodeUtf8("66eI7KeA66eJIOyImOyglQ=="),
				DecodeUtf8("7IiY7KCV7J6Q")
			};
		}
		return new List<string> { "Enabled", "RootFolder", "RvtFileName", "RelativePath", "CentralPath", "Discipline", "TrackElementChanges", "BlockFamilyLoadEdit", "BlockTypeChanges", "BlockNestedOnlyStandalonePlacement", "LastUpdatedUtc", "LastUpdatedBy" };
	}

	private static List<string> BuildTemplateRow(string enabled, string applyFolder, string rvtFileName, string userName, string loadEditFamily, string addDeleteType, string notes)
	{
		return new List<string>
		{
			enabled,
			applyFolder ?? string.Empty,
			rvtFileName ?? string.Empty,
			userName,
			loadEditFamily,
			addDeleteType,
			notes
		};
	}

	private static List<string> BuildRvtPolicyRow(string sourceFolder, string rvtPath, bool korean)
	{
		string relativePath = MakeRelativePath(sourceFolder, rvtPath);
		return new List<string>
		{
			"O",
			sourceFolder,
			Path.GetFileName(rvtPath),
			"*",
			"X",
			"X",
			(korean ? "RVT 폴더에서 생성: " : "Generated from RVT folder: ") + relativePath
		};
	}

	private static List<string> BuildFileGuardRow(FamilyBrowserFileGuardPolicy fileGuard, FamilyBrowserFileGuardTarget target)
	{
		return new List<string>
		{
			BoolToken(target?.Enabled ?? false),
			(fileGuard == null) ? string.Empty : (fileGuard.RootFolder ?? string.Empty),
			(target == null) ? string.Empty : (target.FileName ?? string.Empty),
			(target == null) ? string.Empty : (target.RelativePath ?? string.Empty),
			(target == null) ? string.Empty : (target.CentralPath ?? string.Empty),
			(target == null) ? string.Empty : (target.Discipline ?? string.Empty),
			BoolToken(target?.TrackElementChanges ?? true),
			BoolToken(target?.BlockFamilyLoadAndEdit ?? false),
			BoolToken(target?.BlockTypeChanges ?? false),
			BoolToken(target?.BlockNestedOnlyStandalonePlacement ?? false),
			(target == null) ? string.Empty : (target.LastUpdatedUtc ?? string.Empty),
			(target == null) ? string.Empty : (target.LastUpdatedBy ?? string.Empty)
		};
	}

	private static string BoolToken(bool value)
	{
		if (!value)
		{
			return "X";
		}
		return "O";
	}

	private static string DecodeUtf8(string base64Text)
	{
		return Encoding.UTF8.GetString(Convert.FromBase64String(base64Text));
	}

	private static bool IsLikelyRevitBackup(string filePath)
	{
		string name = Path.GetFileNameWithoutExtension(filePath);
		int dotIndex = name?.LastIndexOf('.') ?? (-1);
		checked
		{
			if (dotIndex < 0 || dotIndex >= name.Length - 1)
			{
				return false;
			}
			string suffix = name.Substring(dotIndex + 1);
			return suffix.Length == 4 && suffix.All([SpecialName] (char ch) => char.IsDigit(ch));
		}
	}

	private static string MakeRelativePath(string root, string filePath)
	{
		string MakeRelativePath;
		try
		{
			Uri uri = new Uri(root.EndsWith(Path.DirectorySeparatorChar.ToString(), StringComparison.Ordinal) ? root : (root + Conversions.ToString(Path.DirectorySeparatorChar)));
			Uri pathUri = new Uri(filePath);
			MakeRelativePath = Uri.UnescapeDataString(uri.MakeRelativeUri(pathUri).ToString()).Replace('/', Path.DirectorySeparatorChar);
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			MakeRelativePath = filePath;
			ProjectData.ClearProjectError();
		}
		return MakeRelativePath;
	}

	private static void WriteWorkbook(string outputPath, string sheetName, List<string> headers, List<List<string>> rows)
	{
		//IL_0052: Unknown result type (might be due to invalid IL or missing references)
		//IL_0058: Expected O, but got Unknown
		if (string.IsNullOrWhiteSpace(outputPath))
		{
			throw new ArgumentException(FamilyBrowserLanguageService.Text("Output Excel path is empty.", "내보낼 Excel 경로가 비어 있습니다."), "outputPath");
		}
		string outputFolder = Path.GetDirectoryName(outputPath);
		if (!string.IsNullOrWhiteSpace(outputFolder))
		{
			Directory.CreateDirectory(outputFolder);
		}
		if (File.Exists(outputPath))
		{
			File.Delete(outputPath);
		}
		using FileStream stream = new FileStream(outputPath, FileMode.CreateNew, FileAccess.ReadWrite, FileShare.None);
		ZipArchive archive = new ZipArchive((Stream)stream, (ZipArchiveMode)1);
		try
		{
			AddEntry(archive, "[Content_Types].xml", BuildContentTypesXml());
			AddEntry(archive, "_rels/.rels", BuildRootRelationshipsXml());
			AddEntry(archive, "docProps/app.xml", BuildAppXml());
			AddEntry(archive, "docProps/core.xml", BuildCoreXml());
			AddEntry(archive, "xl/workbook.xml", BuildWorkbookXml(sheetName));
			AddEntry(archive, "xl/_rels/workbook.xml.rels", BuildWorkbookRelationshipsXml());
			AddEntry(archive, "xl/styles.xml", BuildStylesXml());
			AddEntry(archive, "xl/worksheets/sheet1.xml", BuildWorksheetXml(headers, rows));
		}
		finally
		{
			((IDisposable)archive)?.Dispose();
		}
	}

	private static void AddEntry(ZipArchive archive, string entryName, string content)
	{
		ZipArchiveEntry entry = archive.CreateEntry(entryName, CompressionLevel.Optimal);
		using StreamWriter writer = new StreamWriter(entry.Open(), new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));
		writer.Write(content);
	}

	private static string BuildContentTypesXml()
	{
		return "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"><Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/><Default Extension=\"xml\" ContentType=\"application/xml\"/><Override PartName=\"/docProps/app.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.extended-properties+xml\"/><Override PartName=\"/docProps/core.xml\" ContentType=\"application/vnd.openxmlformats-package.core-properties+xml\"/><Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/><Override PartName=\"/xl/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\"/><Override PartName=\"/xl/worksheets/sheet1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/></Types>";
	}

	private static string BuildRootRelationshipsXml()
	{
		return "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/><Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties\" Target=\"docProps/core.xml\"/><Relationship Id=\"rId3\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties\" Target=\"docProps/app.xml\"/></Relationships>";
	}

	private static string BuildWorkbookXml(string sheetName)
	{
		return "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><sheets><sheet name=\"" + XmlEscape(string.IsNullOrWhiteSpace(sheetName) ? "Policy" : sheetName) + "\" sheetId=\"1\" r:id=\"rId1\"/></sheets></workbook>";
	}

	private static string BuildWorkbookRelationshipsXml()
	{
		return "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet1.xml\"/><Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/></Relationships>";
	}

	private static string BuildStylesXml()
	{
		return "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><fonts count=\"2\"><font><sz val=\"11\"/><name val=\"Segoe UI\"/></font><font><b/><sz val=\"11\"/><name val=\"Segoe UI\"/></font></fonts><fills count=\"2\"><fill><patternFill patternType=\"none\"/></fill><fill><patternFill patternType=\"gray125\"/></fill></fills><borders count=\"1\"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/></cellStyleXfs><cellXfs count=\"2\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\"/><xf numFmtId=\"0\" fontId=\"1\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyFont=\"1\"/></cellXfs><cellStyles count=\"1\"><cellStyle name=\"Normal\" xfId=\"0\" builtinId=\"0\"/></cellStyles></styleSheet>";
	}

	private static string BuildAppXml()
	{
		return "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Properties xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties\" xmlns:vt=\"http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes\"><Application>KKY Family Browser</Application></Properties>";
	}

	private static string BuildCoreXml()
	{
		string stamp = DateTime.UtcNow.ToString("yyyy-MM-ddTHH:mm:ssZ", CultureInfo.InvariantCulture);
		return "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><cp:coreProperties xmlns:cp=\"http://schemas.openxmlformats.org/package/2006/metadata/core-properties\" xmlns:dc=\"http://purl.org/dc/elements/1.1/\" xmlns:dcterms=\"http://purl.org/dc/terms/\" xmlns:dcmitype=\"http://purl.org/dc/dcmitype/\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\"><dc:creator>KKY Family Browser</dc:creator><cp:lastModifiedBy>KKY Family Browser</cp:lastModifiedBy><dcterms:created xsi:type=\"dcterms:W3CDTF\">" + stamp + "</dcterms:created><dcterms:modified xsi:type=\"dcterms:W3CDTF\">" + stamp + "</dcterms:modified></cp:coreProperties>";
	}

	private static string BuildWorksheetXml(List<string> headers, List<List<string>> rows)
	{
		checked
		{
			int rowCount = (rows?.Count ?? 0) + 1;
			int colCount = headers?.Count ?? 0;
			string lastRef = ColumnName(Math.Max(1, colCount)) + Math.Max(1, rowCount).ToString(CultureInfo.InvariantCulture);
			StringBuilder builder = new StringBuilder();
			builder.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
			builder.Append("<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">");
			builder.Append("<dimension ref=\"A1:" + lastRef + "\"/>");
			builder.Append("<sheetViews><sheetView workbookViewId=\"0\"><pane ySplit=\"1\" topLeftCell=\"A2\" activePane=\"bottomLeft\" state=\"frozen\"/></sheetView></sheetViews>");
			builder.Append(BuildColumnsXml(colCount));
			builder.Append("<sheetData>");
			builder.Append(BuildRowXml(1, headers, 1));
			if (rows != null)
			{
				int num = rows.Count - 1;
				for (int i = 0; i <= num; i++)
				{
					builder.Append(BuildRowXml(i + 2, rows[i], 0));
				}
			}
			builder.Append("</sheetData>");
			builder.Append("<autoFilter ref=\"A1:" + lastRef + "\"/>");
			builder.Append("</worksheet>");
			return builder.ToString();
		}
	}

	private static string BuildColumnsXml(int colCount)
	{
		StringBuilder builder = new StringBuilder();
		builder.Append("<cols>");
		for (int i = 1; i <= colCount; i = checked(i + 1))
		{
			int width = i switch
			{
				2 => 52, 
				3 => 34, 
				7 => 52, 
				_ => 18, 
			};
			builder.Append("<col min=\"" + i.ToString(CultureInfo.InvariantCulture) + "\" max=\"" + i.ToString(CultureInfo.InvariantCulture) + "\" width=\"" + width.ToString(CultureInfo.InvariantCulture) + "\" customWidth=\"1\"/>");
		}
		builder.Append("</cols>");
		return builder.ToString();
	}

	private static string BuildRowXml(int rowIndex, IList<string> values, int styleIndex)
	{
		StringBuilder builder = new StringBuilder();
		builder.Append("<row r=\"" + rowIndex.ToString(CultureInfo.InvariantCulture) + "\">");
		checked
		{
			if (values != null)
			{
				int num = values.Count - 1;
				for (int col = 0; col <= num; col++)
				{
					builder.Append(BuildCellXml(rowIndex, col + 1, values[col], styleIndex));
				}
			}
			builder.Append("</row>");
			return builder.ToString();
		}
	}

	private static string BuildCellXml(int rowIndex, int columnIndex, string value, int styleIndex)
	{
		string styleText = ((styleIndex > 0) ? (" s=\"" + styleIndex.ToString(CultureInfo.InvariantCulture) + "\"") : string.Empty);
		return "<c r=\"" + ColumnName(columnIndex) + rowIndex.ToString(CultureInfo.InvariantCulture) + "\" t=\"inlineStr\"" + styleText + "><is><t xml:space=\"preserve\">" + XmlEscape(value) + "</t></is></c>";
	}

	private static string ColumnName(int columnIndex)
	{
		int value = columnIndex;
		StringBuilder builder = new StringBuilder();
		while (value > 0)
		{
			checked
			{
				value--;
				builder.Insert(0, Strings.ChrW(65 + unchecked(value % 26)));
			}
			value /= 26;
		}
		return builder.ToString();
	}

	private static string XmlEscape(string value)
	{
		return SecurityElement.Escape(value ?? string.Empty);
	}
}
