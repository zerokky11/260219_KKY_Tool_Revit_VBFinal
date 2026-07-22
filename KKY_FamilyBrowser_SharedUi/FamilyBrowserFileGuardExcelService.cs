using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Xml.Linq;

public sealed class FamilyBrowserFileGuardExcelImportResult
{
	public FamilyBrowserFileGuardPolicy Policy { get; set; }

	public int ImportedRowCount { get; set; }

	public int SkippedRowCount { get; set; }

	public List<string> Warnings { get; set; }

	public FamilyBrowserFileGuardExcelImportResult()
	{
		Policy = FamilyBrowserFileGuardPolicy.CreateDefault();
		Warnings = new List<string>();
	}
}

public static class FamilyBrowserFileGuardExcelService
{
	private sealed class WorkbookRow
	{
		public int RowNumber { get; set; }

		public Dictionary<string, string> Values { get; set; }

		public WorkbookRow()
		{
			Values = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
		}
	}

	public static FamilyBrowserFileGuardExcelImportResult Import(
		string inputPath,
		FamilyBrowserStandardPolicy standardPolicy,
		string currentUser,
		bool korean)
	{
		if (string.IsNullOrWhiteSpace(inputPath) || !File.Exists(inputPath))
		{
			throw new FileNotFoundException(Text(korean, "The file guard Excel workbook was not found.", "파일별 권한 Excel 파일을 찾을 수 없습니다."), inputPath);
		}
		if (!string.Equals(Path.GetExtension(inputPath), ".xlsx", StringComparison.OrdinalIgnoreCase))
		{
			throw new InvalidDataException(Text(korean, "Only .xlsx workbooks can be imported.", ".xlsx 형식만 불러올 수 있습니다."));
		}

		List<WorkbookRow> rows = ReadRows(inputPath);
		FamilyBrowserFileGuardExcelImportResult result = new FamilyBrowserFileGuardExcelImportResult();
		List<FamilyBrowserFileGuardTarget> targets = new List<FamilyBrowserFileGuardTarget>();
		Dictionary<string, int> targetIndexByPath = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
		string policyRoot = string.Empty;
		string nowText = DateTime.UtcNow.ToString("O", CultureInfo.InvariantCulture);

		foreach (WorkbookRow row in rows)
		{
			string rootFolder = Value(row, "rootfolder", "applyfolder", "적용폴더", "루트폴더");
			string fileName = Value(row, "rvtfilename", "rvtfile", "filename", "rvt파일명", "rvt파일", "파일명");
			string relativePath = Value(row, "relativepath", "상대경로");
			string centralPath = Value(row, "centralpath", "filepath", "modelpath", "중앙경로", "센트럴경로", "파일경로", "경로");
			if (string.IsNullOrWhiteSpace(fileName) && string.IsNullOrWhiteSpace(relativePath) && string.IsNullOrWhiteSpace(centralPath))
			{
				continue;
			}
			if (string.IsNullOrWhiteSpace(policyRoot) && !string.IsNullOrWhiteSpace(rootFolder))
			{
				policyRoot = rootFolder.Trim();
			}

			string resolvedPath = ResolveTargetPath(rootFolder, relativePath, centralPath, fileName);
			if (string.IsNullOrWhiteSpace(resolvedPath))
			{
				result.SkippedRowCount++;
				result.Warnings.Add(RowWarning(row.RowNumber, korean, "No usable RVT path was provided.", "사용할 수 있는 RVT 경로가 없습니다."));
				continue;
			}
			if (string.IsNullOrWhiteSpace(fileName))
			{
				fileName = SafeFileName(resolvedPath);
			}
			if (!string.Equals(Path.GetExtension(fileName), ".rvt", StringComparison.OrdinalIgnoreCase) &&
				!string.Equals(Path.GetExtension(resolvedPath), ".rvt", StringComparison.OrdinalIgnoreCase))
			{
				result.SkippedRowCount++;
				result.Warnings.Add(RowWarning(row.RowNumber, korean, "The row does not identify an RVT file.", "RVT 파일을 가리키는 행이 아닙니다."));
				continue;
			}

			string rawDiscipline = Value(row, "discipline", "trade", "field", "공종", "분야");
			FamilyBrowserStandardLibrarySlot slot = FamilyBrowserFileGuardDisciplineService.ResolveSlot(standardPolicy, rawDiscipline);
			if (slot == null)
			{
				result.SkippedRowCount++;
				result.Warnings.Add(RowWarning(row.RowNumber, korean,
					string.IsNullOrWhiteSpace(rawDiscipline) ? "Discipline is required for every RVT row." : "The discipline is not registered in Standards: " + rawDiscipline,
					string.IsNullOrWhiteSpace(rawDiscipline) ? "모든 RVT 행에 공종을 입력해야 합니다." : "표준 관리에 등록되지 않은 공종입니다: " + rawDiscipline));
				continue;
			}

			FamilyBrowserFileGuardTarget target = new FamilyBrowserFileGuardTarget
			{
				Enabled = BoolValue(row, true, "enabled", "apply", "use", "사용", "적용"),
				FileName = string.IsNullOrWhiteSpace(fileName) ? SafeFileName(resolvedPath) : fileName.Trim(),
				CentralPath = resolvedPath,
				RelativePath = string.IsNullOrWhiteSpace(relativePath) ? MakeRelativePath(policyRoot, resolvedPath) : relativePath.Trim(),
				Discipline = slot.Discipline ?? string.Empty,
				TrackElementChanges = BoolValue(row, true, "trackelementchanges", "trackchanges", "요소생성수정삭제추적", "요소변경추적"),
				TrackElementChangesConfigured = true,
				BlockFamilyLoadAndEdit = BoolValue(row, true, "blockfamilyloadedit", "blockfamilyloadandedit", "패밀리로드편집차단", "패밀리로드편집"),
				BlockTypeChanges = BoolValue(row, true, "blocktypechanges", "타입변경차단", "타입변경"),
				BlockNestedOnlyStandalonePlacement = BoolValue(row, false, "blocknestedonlystandaloneplacement", "blocknestedonly", "하위전용패밀리단독모델링금지", "하위패밀리단독배치금지"),
				LastUpdatedUtc = nowText,
				LastUpdatedBy = currentUser ?? string.Empty
			};
			string key = NormalizePathKey(resolvedPath);
			int existingIndex;
			if (targetIndexByPath.TryGetValue(key, out existingIndex))
			{
				FamilyBrowserFileGuardTarget merged = FamilyBrowserFileGuardPathMatcher.MergeConservativeTargets(new FamilyBrowserFileGuardTarget[] { targets[existingIndex], target });
				targets[existingIndex] = merged;
				result.Warnings.Add(RowWarning(row.RowNumber, korean, "A duplicate RVT path was merged using the strictest guard settings.", "같은 RVT 경로의 권한을 가장 엄격한 차단 설정으로 병합했습니다."));
				if (merged == null || string.IsNullOrWhiteSpace(merged.Discipline))
				{
					result.Warnings.Add(RowWarning(row.RowNumber, korean, "Duplicate rows assign different trades. Select one trade in File Guard before saving.", "중복 행의 공종이 서로 다릅니다. 저장 전에 파일별 권한 화면에서 공종 하나를 선택하세요."));
				}
			}
			else
			{
				targetIndexByPath[key] = targets.Count;
				targets.Add(target);
			}
		}

		result.ImportedRowCount = targets.Count;
		result.Policy = new FamilyBrowserFileGuardPolicy
		{
			Enabled = targets.Any(delegate(FamilyBrowserFileGuardTarget target) { return target != null && target.Enabled; }),
			RootFolder = policyRoot,
			Targets = targets,
			LastUpdatedUtc = nowText,
			LastUpdatedBy = currentUser ?? string.Empty
		};
		return result;
	}

	private static List<WorkbookRow> ReadRows(string path)
	{
		using (ZipArchive archive = ZipFile.OpenRead(path))
		{
			List<string> sharedStrings = ReadSharedStrings(archive);
			ZipArchiveEntry sheetEntry = archive.GetEntry("xl/worksheets/sheet1.xml") ??
				archive.Entries.FirstOrDefault(delegate(ZipArchiveEntry entry)
				{
					return entry.FullName.StartsWith("xl/worksheets/", StringComparison.OrdinalIgnoreCase) &&
						entry.FullName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase);
				});
			if (sheetEntry == null)
			{
				throw new InvalidDataException("Worksheet XML was not found in the file guard workbook.");
			}
			List<List<string>> table = new List<List<string>>();
			using (Stream stream = sheetEntry.Open())
			{
				XDocument document = XDocument.Load(stream);
				XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
				foreach (XElement rowElement in document.Descendants(ns + "row"))
				{
					Dictionary<int, string> cells = new Dictionary<int, string>();
					foreach (XElement cell in rowElement.Elements(ns + "c"))
					{
						int column = ResolveColumnIndex((string)cell.Attribute("r"));
						if (column <= 0)
						{
							column = cells.Count + 1;
						}
						cells[column] = ReadCellValue(cell, ns, sharedStrings);
					}
					if (cells.Count > 0)
					{
						int maxColumn = cells.Keys.Max();
						List<string> values = new List<string>();
						for (int column = 1; column <= maxColumn; column++)
						{
							values.Add(cells.ContainsKey(column) ? cells[column] : string.Empty);
						}
						table.Add(values);
					}
				}
			}
			if (table.Count == 0)
			{
				return new List<WorkbookRow>();
			}
			int headerIndex = table.FindIndex(delegate(List<string> values)
			{
				return values.Any(delegate(string value) { return !string.IsNullOrWhiteSpace(value); });
			});
			if (headerIndex < 0)
			{
				return new List<WorkbookRow>();
			}
			List<string> headers = table[headerIndex].Select(NormalizeHeader).ToList();
			List<WorkbookRow> result = new List<WorkbookRow>();
			for (int rowIndex = headerIndex + 1; rowIndex < table.Count; rowIndex++)
			{
				List<string> values = table[rowIndex];
				if (values.All(delegate(string value) { return string.IsNullOrWhiteSpace(value); }))
				{
					continue;
				}
				WorkbookRow row = new WorkbookRow { RowNumber = rowIndex + 1 };
				for (int column = 0; column < values.Count && column < headers.Count; column++)
				{
					if (!string.IsNullOrWhiteSpace(headers[column]))
					{
						row.Values[headers[column]] = values[column] ?? string.Empty;
					}
				}
				result.Add(row);
			}
			return result;
		}
	}

	private static List<string> ReadSharedStrings(ZipArchive archive)
	{
		List<string> result = new List<string>();
		ZipArchiveEntry entry = archive.GetEntry("xl/sharedStrings.xml");
		if (entry == null)
		{
			return result;
		}
		using (Stream stream = entry.Open())
		{
			XDocument document = XDocument.Load(stream);
			XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
			foreach (XElement item in document.Descendants(ns + "si"))
			{
				result.Add(string.Concat(item.Descendants(ns + "t").Select(delegate(XElement text) { return text.Value; })));
			}
		}
		return result;
	}

	private static string ReadCellValue(XElement cell, XNamespace ns, List<string> sharedStrings)
	{
		string type = (string)cell.Attribute("t") ?? string.Empty;
		if (string.Equals(type, "inlineStr", StringComparison.OrdinalIgnoreCase))
		{
			return string.Concat(cell.Descendants(ns + "t").Select(delegate(XElement text) { return text.Value; }));
		}
		string raw = (string)cell.Element(ns + "v") ?? string.Empty;
		int index;
		if (string.Equals(type, "s", StringComparison.OrdinalIgnoreCase) &&
			int.TryParse(raw, NumberStyles.Integer, CultureInfo.InvariantCulture, out index) &&
			index >= 0 && index < sharedStrings.Count)
		{
			return sharedStrings[index];
		}
		return raw;
	}

	private static int ResolveColumnIndex(string cellReference)
	{
		string letters = new string((cellReference ?? string.Empty).TakeWhile(delegate(char value) { return char.IsLetter(value); }).ToArray()).ToUpperInvariant();
		int result = 0;
		foreach (char value in letters)
		{
			result = checked(result * 26 + value - 'A' + 1);
		}
		return result;
	}

	private static string Value(WorkbookRow row, params string[] aliases)
	{
		if (row == null || row.Values == null)
		{
			return string.Empty;
		}
		foreach (string alias in aliases ?? new string[0])
		{
			string value;
			if (row.Values.TryGetValue(NormalizeHeader(alias), out value))
			{
				return (value ?? string.Empty).Trim();
			}
		}
		return string.Empty;
	}

	private static bool BoolValue(WorkbookRow row, bool defaultValue, params string[] aliases)
	{
		string value = Value(row, aliases);
		if (string.IsNullOrWhiteSpace(value))
		{
			return defaultValue;
		}
		string normalized = NormalizeHeader(value);
		if (normalized == "o" || normalized == "1" || normalized == "true" || normalized == "yes" || normalized == "y" || normalized == "사용" || normalized == "적용")
		{
			return true;
		}
		if (normalized == "x" || normalized == "0" || normalized == "false" || normalized == "no" || normalized == "n" || normalized == "미사용" || normalized == "해제")
		{
			return false;
		}
		return defaultValue;
	}

	private static string ResolveTargetPath(string rootFolder, string relativePath, string centralPath, string fileName)
	{
		if (!string.IsNullOrWhiteSpace(centralPath))
		{
			return NormalizePath(centralPath);
		}
		string child = !string.IsNullOrWhiteSpace(relativePath) ? relativePath : fileName;
		if (!string.IsNullOrWhiteSpace(rootFolder) && !string.IsNullOrWhiteSpace(child))
		{
			try
			{
				return NormalizePath(Path.Combine(rootFolder, child));
			}
			catch
			{
			}
		}
		return NormalizePath(child);
	}

	private static string MakeRelativePath(string rootFolder, string filePath)
	{
		if (string.IsNullOrWhiteSpace(rootFolder) || string.IsNullOrWhiteSpace(filePath))
		{
			return SafeFileName(filePath);
		}
		try
		{
			string root = Path.GetFullPath(rootFolder).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar) + Path.DirectorySeparatorChar;
			Uri rootUri = new Uri(root);
			Uri fileUri = new Uri(Path.GetFullPath(filePath));
			return Uri.UnescapeDataString(rootUri.MakeRelativeUri(fileUri).ToString()).Replace('/', Path.DirectorySeparatorChar);
		}
		catch
		{
			return SafeFileName(filePath);
		}
	}

	private static string SafeFileName(string path)
	{
		try
		{
			return Path.GetFileName(path ?? string.Empty) ?? string.Empty;
		}
		catch
		{
			return path ?? string.Empty;
		}
	}

	private static string NormalizePath(string value)
	{
		string text = Environment.ExpandEnvironmentVariables((value ?? string.Empty).Trim()).Replace('/', '\\');
		if (text.Length == 0)
		{
			return string.Empty;
		}
		try
		{
			if (Path.IsPathRooted(text))
			{
				return Path.GetFullPath(text);
			}
		}
		catch
		{
		}
		return text;
	}

	private static string NormalizePathKey(string value)
	{
		return FamilyBrowserFileGuardPathMatcher.BuildStablePolicyPathKey(value);
	}

	private static string NormalizeHeader(string value)
	{
		StringBuilder builder = new StringBuilder();
		foreach (char character in (value ?? string.Empty).Trim().ToLowerInvariant())
		{
			if (char.IsLetterOrDigit(character) || character >= 0xAC00)
			{
				builder.Append(character);
			}
		}
		return builder.ToString();
	}

	private static string RowWarning(int rowNumber, bool korean, string english, string koreanText)
	{
		return Text(korean, "Row " + rowNumber.ToString(CultureInfo.InvariantCulture) + ": " + english, rowNumber.ToString(CultureInfo.InvariantCulture) + "행: " + koreanText);
	}

	private static string Text(bool korean, string english, string koreanText)
	{
		return korean ? koreanText : english;
	}
}
