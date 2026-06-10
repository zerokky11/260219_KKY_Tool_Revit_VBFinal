using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading;
using System.Xml.Linq;
using Microsoft.VisualBasic.CompilerServices;

public sealed class FamilyBrowserPermissionExcelPolicyService
{
	[Serializable]
	[CompilerGenerated]
	internal sealed class _Closure_0024__
	{
		public static readonly _Closure_0024__ _0024I;

		public static Func<FamilyBrowserPermissionExcelRow, bool> _0024I15_002D0;

		public static Func<FamilyBrowserPermissionExcelRow, bool> _0024I17_002D0;

		public static Func<string, bool> _0024I19_002D0;

		public static Func<string, bool> _0024I20_002D1;

		public static Func<List<string>, bool> _0024I20_002D0;

		public static Func<XElement, string> _0024I21_002D0;

		public static Func<XElement, string> _0024I24_002D0;

		public static Func<char, bool> _0024I25_002D0;

		public static Func<string, bool> _0024I26_002D0;

		static _Closure_0024__()
		{
			_0024I = new _Closure_0024__();
		}

		[SpecialName]
		internal bool _Lambda_0024__15_002D0(FamilyBrowserPermissionExcelRow row)
		{
			return IsRowEnabled(row);
		}

		[SpecialName]
		internal bool _Lambda_0024__17_002D0(FamilyBrowserPermissionExcelRow row)
		{
			return IsRowEnabled(row);
		}

		[SpecialName]
		internal bool _Lambda_0024__19_002D0(string x)
		{
			return x != null;
		}

		[SpecialName]
		internal bool _Lambda_0024__20_002D0(List<string> x)
		{
			return x.Any([SpecialName] (string v) => !string.IsNullOrWhiteSpace(v));
		}

		[SpecialName]
		internal bool _Lambda_0024__20_002D1(string v)
		{
			return !string.IsNullOrWhiteSpace(v);
		}

		[SpecialName]
		internal string _Lambda_0024__21_002D0(XElement x)
		{
			return x.Value;
		}

		[SpecialName]
		internal string _Lambda_0024__24_002D0(XElement x)
		{
			return x.Value;
		}

		[SpecialName]
		internal bool _Lambda_0024__25_002D0(char ch)
		{
			return char.IsLetter(ch);
		}

		[SpecialName]
		internal bool _Lambda_0024__26_002D0(string x)
		{
			return string.IsNullOrWhiteSpace(x);
		}
	}

	[CompilerGenerated]
	internal sealed class _Closure_0024__15_002D0
	{
		public FamilyBrowserProjectPolicyContext _0024VB_0024Local_context;

		public string _0024VB_0024Local_currentUser;

		public _Closure_0024__15_002D0(_Closure_0024__15_002D0 arg0)
		{
			if (arg0 != null)
			{
				_0024VB_0024Local_context = arg0._0024VB_0024Local_context;
				_0024VB_0024Local_currentUser = arg0._0024VB_0024Local_currentUser;
			}
		}

		[SpecialName]
		internal bool _Lambda_0024__1(FamilyBrowserPermissionExcelRow row)
		{
			return MatchesProject(row, _0024VB_0024Local_context);
		}

		[SpecialName]
		internal bool _Lambda_0024__2(FamilyBrowserPermissionExcelRow row)
		{
			return MatchesUser(row, _0024VB_0024Local_currentUser);
		}

		[SpecialName]
		internal bool _Lambda_0024__3(FamilyBrowserPermissionExcelRow row)
		{
			if (MatchesProject(row, _0024VB_0024Local_context))
			{
				return MatchesUser(row, _0024VB_0024Local_currentUser);
			}
			return false;
		}
	}

	[CompilerGenerated]
	internal sealed class _Closure_0024__22_002D0
	{
		public string _0024VB_0024Local_requestedSheetName;

		public _Closure_0024__22_002D0(_Closure_0024__22_002D0 arg0)
		{
			if (arg0 != null)
			{
				_0024VB_0024Local_requestedSheetName = arg0._0024VB_0024Local_requestedSheetName;
			}
		}

		[SpecialName]
		internal bool _Lambda_0024__0(XElement sheet)
		{
			if (!string.IsNullOrWhiteSpace(_0024VB_0024Local_requestedSheetName))
			{
				return string.Equals((sheet.Attribute("name") == null) ? string.Empty : sheet.Attribute("name").Value, _0024VB_0024Local_requestedSheetName, StringComparison.OrdinalIgnoreCase);
			}
			return true;
		}
	}

	private static readonly object CacheLock = RuntimeHelpers.GetObjectValue(new object());

	private static string CachedPath = string.Empty;

	private static string CachedSheetName = string.Empty;

	private static long CachedWriteTicks = -1L;

	private static List<FamilyBrowserPermissionExcelRow> CachedRows = new List<FamilyBrowserPermissionExcelRow>();

	private static string CachedLoadedUtc = string.Empty;

	private static string CachedError = string.Empty;

	private static DateTime CachedStatUtc = DateTime.MinValue;

	private const double FileStatCacheSeconds = 10.0;

	private FamilyBrowserPermissionExcelPolicyService()
	{
	}

	public static void ClearCache()
	{
		object cacheLock = CacheLock;
		ObjectFlowControl.CheckForSyncLockOnValueType(cacheLock);
		bool lockTaken = false;
		try
		{
			Monitor.Enter(cacheLock, ref lockTaken);
			CachedPath = string.Empty;
			CachedSheetName = string.Empty;
			CachedWriteTicks = -1L;
			CachedRows = new List<FamilyBrowserPermissionExcelRow>();
			CachedLoadedUtc = string.Empty;
			CachedError = string.Empty;
			CachedStatUtc = DateTime.MinValue;
		}
		finally
		{
			if (lockTaken)
			{
				Monitor.Exit(cacheLock);
			}
		}
	}

	public static string ResolveRole(FamilyBrowserStandardPolicy policy, string currentUser, FamilyBrowserProjectPolicyContext context)
	{
		return string.Empty;
	}

	public static FamilyBrowserPermissionExcelDecision ResolvePermission(FamilyBrowserStandardPolicy policy, string currentUser, string permission, FamilyBrowserProjectPolicyContext context)
	{
		FamilyBrowserPermissionExcelDecision decision = new FamilyBrowserPermissionExcelDecision();
		if (!IsNativeGuardPermission(permission))
		{
			return decision;
		}
		try
		{
			FamilyBrowserPermissionExcelRow row = FindMatchingRow(policy, currentUser, context);
			if (row == null)
			{
				return decision;
			}
			string token = ResolvePermissionToken(row, permission);
			if (string.IsNullOrWhiteSpace(token))
			{
				return decision;
			}
			bool parsed = default(bool);
			if (!TryParseYesNo(token, ref parsed))
			{
				return decision;
			}
			decision.HasDecision = true;
			decision.Allowed = parsed;
			decision.Role = NormalizeRole(row.Role);
			decision.SourcePath = ResolveExcelPath(policy);
			decision.SourceRow = row.RowNumber;
			decision.Message = "Excel permission row " + row.RowNumber.ToString(CultureInfo.InvariantCulture);
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ProjectData.ClearProjectError();
		}
		return decision;
	}

	public static FamilyBrowserPermissionExcelStatus BuildStatus(FamilyBrowserStandardPolicy policy)
	{
		FamilyBrowserPermissionExcelSettings settings = policy?.PermissionExcel;
		FamilyBrowserPermissionExcelStatus status = new FamilyBrowserPermissionExcelStatus
		{
			Enabled = (settings?.Enabled ?? false),
			Path = ResolveExcelPath(policy),
			SheetName = ((settings == null) ? string.Empty : settings.SheetName)
		};
		if (!status.Enabled || string.IsNullOrWhiteSpace(status.Path))
		{
			return status;
		}
		status.Exists = File.Exists(status.Path);
		if (!status.Exists)
		{
			status.LastError = "Excel permission file was not found.";
			return status;
		}
		try
		{
			List<FamilyBrowserPermissionExcelRow> rows = LoadRows(policy);
			status.RowCount = rows.Count;
			status.LastLoadedUtc = CachedLoadedUtc;
			status.LastError = CachedError;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			status.LastError = ex2.Message;
			ProjectData.ClearProjectError();
		}
		return status;
	}

	public static FamilyBrowserPermissionExcelDiagnostic BuildDiagnostic(FamilyBrowserStandardPolicy policy, string currentUser, FamilyBrowserProjectPolicyContext context)
	{
		_Closure_0024__15_002D0 arg = default(_Closure_0024__15_002D0);
		_Closure_0024__15_002D0 CS_0024_003C_003E8__locals15 = new _Closure_0024__15_002D0(arg);
		CS_0024_003C_003E8__locals15._0024VB_0024Local_currentUser = currentUser;
		CS_0024_003C_003E8__locals15._0024VB_0024Local_context = context;
		FamilyBrowserPermissionExcelDiagnostic diagnostic = new FamilyBrowserPermissionExcelDiagnostic
		{
			CurrentUser = (CS_0024_003C_003E8__locals15._0024VB_0024Local_currentUser ?? string.Empty),
			ProjectTitle = ((CS_0024_003C_003E8__locals15._0024VB_0024Local_context == null) ? string.Empty : CS_0024_003C_003E8__locals15._0024VB_0024Local_context.ProjectTitle),
			ModelPath = ((CS_0024_003C_003E8__locals15._0024VB_0024Local_context == null) ? string.Empty : CS_0024_003C_003E8__locals15._0024VB_0024Local_context.ModelPath),
			CentralPath = ((CS_0024_003C_003E8__locals15._0024VB_0024Local_context == null) ? string.Empty : CS_0024_003C_003E8__locals15._0024VB_0024Local_context.CentralPath),
			StandardTarget = ((CS_0024_003C_003E8__locals15._0024VB_0024Local_context == null) ? string.Empty : CS_0024_003C_003E8__locals15._0024VB_0024Local_context.StandardTarget),
			SourcePath = ResolveExcelPath(policy)
		};
		FamilyBrowserPermissionExcelSettings settings = policy?.PermissionExcel;
		diagnostic.Enabled = settings?.Enabled ?? false;
		diagnostic.SheetName = ((settings == null) ? string.Empty : settings.SheetName);
		diagnostic.Exists = !string.IsNullOrWhiteSpace(diagnostic.SourcePath) && File.Exists(diagnostic.SourcePath);
		if (!diagnostic.Enabled)
		{
			diagnostic.Message = "Excel permission policy is not connected.";
			return diagnostic;
		}
		if (!diagnostic.Exists)
		{
			diagnostic.LastError = "Excel permission file was not found.";
			diagnostic.Message = diagnostic.LastError;
			return diagnostic;
		}
		try
		{
			List<FamilyBrowserPermissionExcelRow> rows = LoadRows(policy);
			diagnostic.RowCount = rows?.Count ?? 0;
			if (rows == null || rows.Count == 0)
			{
				diagnostic.Message = "Excel file is readable, but no usable permission row was found.";
				return diagnostic;
			}
			List<FamilyBrowserPermissionExcelRow> activeRows = rows.Where([SpecialName] (FamilyBrowserPermissionExcelRow row) => IsRowEnabled(row)).ToList();
			diagnostic.ActiveRowCount = activeRows.Count;
			diagnostic.ProjectMatchedRowCount = activeRows.Where([SpecialName] (FamilyBrowserPermissionExcelRow row) => MatchesProject(row, CS_0024_003C_003E8__locals15._0024VB_0024Local_context)).Count();
			diagnostic.UserMatchedRowCount = activeRows.Where([SpecialName] (FamilyBrowserPermissionExcelRow row) => MatchesUser(row, CS_0024_003C_003E8__locals15._0024VB_0024Local_currentUser)).Count();
			FamilyBrowserPermissionExcelRow matchedRow = activeRows.FirstOrDefault([SpecialName] (FamilyBrowserPermissionExcelRow row) => MatchesProject(row, CS_0024_003C_003E8__locals15._0024VB_0024Local_context) && MatchesUser(row, CS_0024_003C_003E8__locals15._0024VB_0024Local_currentUser));
			if (matchedRow == null)
			{
				diagnostic.Message = "No Excel row matched both the current model and current Windows user.";
				return diagnostic;
			}
			diagnostic.Matched = true;
			diagnostic.MatchedRowNumber = matchedRow.RowNumber;
			diagnostic.MatchedRole = string.Empty;
			diagnostic.MatchedUser = matchedRow.UserOrGroup;
			diagnostic.MatchedProjectName = matchedRow.ProjectName;
			diagnostic.MatchedDiscipline = matchedRow.Discipline;
			if (!string.IsNullOrWhiteSpace(matchedRow.ApplyFolder) || !string.IsNullOrWhiteSpace(matchedRow.RvtFileName))
			{
				diagnostic.MatchedMode = "ApplyFolder/RvtFileName";
				diagnostic.MatchedValue = (matchedRow.ApplyFolder ?? string.Empty) + " | " + (matchedRow.RvtFileName ?? string.Empty);
			}
			else
			{
				diagnostic.MatchedMode = matchedRow.MatchMode;
				diagnostic.MatchedValue = matchedRow.MatchValue;
			}
			if (matchedRow.Permissions != null)
			{
				foreach (KeyValuePair<string, string> pair in matchedRow.Permissions)
				{
					diagnostic.PermissionTokens[pair.Key] = pair.Value;
				}
			}
			diagnostic.Message = "Excel permission row " + matchedRow.RowNumber.ToString(CultureInfo.InvariantCulture) + " is applied.";
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			diagnostic.LastError = ex2.Message;
			diagnostic.Message = ex2.Message;
			ProjectData.ClearProjectError();
		}
		return diagnostic;
	}

	public static FamilyBrowserPermissionExcelDecision ResolvePermissionFromDiagnostic(FamilyBrowserPermissionExcelDiagnostic diagnostic, string permission)
	{
		FamilyBrowserPermissionExcelDecision decision = new FamilyBrowserPermissionExcelDecision();
		if (!IsNativeGuardPermission(permission))
		{
			return decision;
		}
		if (diagnostic == null || !diagnostic.Matched)
		{
			return decision;
		}
		try
		{
			string token = string.Empty;
			if (diagnostic.PermissionTokens != null && diagnostic.PermissionTokens.ContainsKey(permission))
			{
				token = diagnostic.PermissionTokens[permission];
			}
			if (string.IsNullOrWhiteSpace(token))
			{
				return decision;
			}
			bool parsed = default(bool);
			if (!TryParseYesNo(token, ref parsed))
			{
				return decision;
			}
			decision.HasDecision = true;
			decision.Allowed = parsed;
			decision.Role = NormalizeRole(diagnostic.MatchedRole);
			if (string.IsNullOrWhiteSpace(decision.Role))
			{
				decision.Role = diagnostic.MatchedRole;
			}
			decision.SourcePath = diagnostic.SourcePath;
			decision.SourceRow = diagnostic.MatchedRowNumber;
			decision.Message = "Excel permission row " + diagnostic.MatchedRowNumber.ToString(CultureInfo.InvariantCulture);
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ProjectData.ClearProjectError();
		}
		return decision;
	}

	private static FamilyBrowserPermissionExcelRow FindMatchingRow(FamilyBrowserStandardPolicy policy, string currentUser, FamilyBrowserProjectPolicyContext context)
	{
		List<FamilyBrowserPermissionExcelRow> rows = LoadRows(policy);
		if (rows == null || rows.Count == 0)
		{
			return null;
		}
		return rows.Where([SpecialName] (FamilyBrowserPermissionExcelRow row) => IsRowEnabled(row)).FirstOrDefault([SpecialName] (FamilyBrowserPermissionExcelRow row) => MatchesProject(row, context) && MatchesUser(row, currentUser));
	}

	private static List<FamilyBrowserPermissionExcelRow> LoadRows(FamilyBrowserStandardPolicy policy)
	{
		FamilyBrowserPermissionExcelSettings settings = policy?.PermissionExcel;
		if (settings == null || !settings.Enabled)
		{
			return new List<FamilyBrowserPermissionExcelRow>();
		}
		string path = ResolveExcelPath(policy);
		if (string.IsNullOrWhiteSpace(path))
		{
			return new List<FamilyBrowserPermissionExcelRow>();
		}
		string sheetName = (settings.SheetName ?? string.Empty).Trim();
		DateTime now = DateTime.UtcNow;
		object cacheLock = CacheLock;
		ObjectFlowControl.CheckForSyncLockOnValueType(cacheLock);
		bool lockTaken = false;
		try
		{
			Monitor.Enter(cacheLock, ref lockTaken);
			if (string.Equals(CachedPath, path, StringComparison.OrdinalIgnoreCase) && string.Equals(CachedSheetName, sheetName, StringComparison.OrdinalIgnoreCase) && CachedWriteTicks >= 0 && (now - CachedStatUtc).TotalSeconds < 10.0)
			{
				return CachedRows;
			}
		}
		finally
		{
			if (lockTaken)
			{
				Monitor.Exit(cacheLock);
			}
		}
		if (!File.Exists(path))
		{
			throw new FileNotFoundException(FamilyBrowserLanguageService.Text("Excel permission file was not found.", "권한 Excel 파일을 찾지 못했습니다."), path);
		}
		long writeTicks = File.GetLastWriteTimeUtc(path).Ticks;
		object cacheLock2 = CacheLock;
		ObjectFlowControl.CheckForSyncLockOnValueType(cacheLock2);
		bool lockTaken2 = false;
		try
		{
			Monitor.Enter(cacheLock2, ref lockTaken2);
			if (string.Equals(CachedPath, path, StringComparison.OrdinalIgnoreCase) && string.Equals(CachedSheetName, sheetName, StringComparison.OrdinalIgnoreCase) && CachedWriteTicks == writeTicks)
			{
				CachedStatUtc = now;
				return CachedRows;
			}
			List<FamilyBrowserPermissionExcelRow> cachedRows = (string.Equals(Path.GetExtension(path), ".csv", StringComparison.OrdinalIgnoreCase) ? ReadCsvRows(path) : ReadXlsxRows(path, sheetName));
			CachedPath = path;
			CachedSheetName = sheetName;
			CachedWriteTicks = writeTicks;
			CachedRows = cachedRows;
			CachedLoadedUtc = DateTime.UtcNow.ToString("O", CultureInfo.InvariantCulture);
			CachedError = string.Empty;
			CachedStatUtc = now;
			return CachedRows;
		}
		finally
		{
			if (lockTaken2)
			{
				Monitor.Exit(cacheLock2);
			}
		}
	}

	private static List<FamilyBrowserPermissionExcelRow> ReadCsvRows(string path)
	{
		List<string> lines = (from x in File.ReadAllLines(path, Encoding.UTF8)
			where x != null
			select x).ToList();
		if (lines.Count == 0)
		{
			return new List<FamilyBrowserPermissionExcelRow>();
		}
		List<string> header = SplitCsvLine(lines[0]);
		List<List<string>> rows = new List<List<string>>();
		checked
		{
			int num = lines.Count - 1;
			for (int i = 1; i <= num; i++)
			{
				rows.Add(SplitCsvLine(lines[i]));
			}
			return MapRows(header, rows, 2);
		}
	}

	private static List<FamilyBrowserPermissionExcelRow> ReadXlsxRows(string path, string sheetName)
	{
		ZipArchive archive = ZipFile.OpenRead(path);
		checked
		{
			try
			{
				List<string> sharedStrings = ReadSharedStrings(archive);
				string sheetEntryName = ResolveWorksheetEntryName(archive, sheetName);
				if (string.IsNullOrWhiteSpace(sheetEntryName))
				{
					throw new InvalidDataException("No worksheet was found in the Excel permission file.");
				}
				ZipArchiveEntry obj = archive.GetEntry(sheetEntryName) ?? throw new InvalidDataException("Worksheet XML was not found: " + sheetEntryName);
				List<List<string>> tableRows = new List<List<string>>();
				using (Stream stream = obj.Open())
				{
					XDocument document = XDocument.Load(stream);
					XNamespace ns = XNamespace.Get("http://schemas.openxmlformats.org/spreadsheetml/2006/main");
					foreach (XElement rowElement in document.Descendants(ns + "sheetData").Elements(ns + "row"))
					{
						Dictionary<int, string> values = new Dictionary<int, string>();
						foreach (XElement cell in rowElement.Elements(ns + "c"))
						{
							int index = ResolveCellColumnIndex((cell.Attribute("r") == null) ? string.Empty : cell.Attribute("r").Value);
							if (index <= 0)
							{
								index = values.Count + 1;
							}
							values[index] = ReadCellValue(cell, ns, sharedStrings);
						}
						if (values.Count > 0)
						{
							int num = values.Keys.Max();
							List<string> rowValues = new List<string>();
							int num2 = num;
							for (int col = 1; col <= num2; col++)
							{
								rowValues.Add(values.ContainsKey(col) ? values[col] : string.Empty);
							}
							tableRows.Add(rowValues);
						}
					}
				}
				List<string> header = tableRows.FirstOrDefault([SpecialName] (List<string> x) => x.Any([SpecialName] (string v) => !string.IsNullOrWhiteSpace(v)));
				if (header == null)
				{
					return new List<FamilyBrowserPermissionExcelRow>();
				}
				int headerIndex = tableRows.IndexOf(header);
				List<List<string>> body = tableRows.Skip(headerIndex + 1).ToList();
				return MapRows(header, body, headerIndex + 2);
			}
			finally
			{
				((IDisposable)archive)?.Dispose();
			}
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
			XNamespace ns = XNamespace.Get("http://schemas.openxmlformats.org/spreadsheetml/2006/main");
			foreach (XElement item in document.Descendants(ns + "si"))
			{
				string text = string.Concat(item.Descendants(ns + "t").Select((_Closure_0024__._0024I21_002D0 == null) ? (_Closure_0024__._0024I21_002D0 = [SpecialName] (XElement x) => x.Value) : _Closure_0024__._0024I21_002D0));
				result.Add(text);
			}
		}
		return result;
	}

	private static string ResolveWorksheetEntryName(ZipArchive archive, string requestedSheetName)
	{
		_Closure_0024__22_002D0 arg = default(_Closure_0024__22_002D0);
		_Closure_0024__22_002D0 CS_0024_003C_003E8__locals3 = new _Closure_0024__22_002D0(arg);
		CS_0024_003C_003E8__locals3._0024VB_0024Local_requestedSheetName = requestedSheetName;
		ZipArchiveEntry workbookEntry = archive.GetEntry("xl/workbook.xml");
		ZipArchiveEntry relsEntry = archive.GetEntry("xl/_rels/workbook.xml.rels");
		if (workbookEntry == null || relsEntry == null)
		{
			return "xl/worksheets/sheet1.xml";
		}
		Dictionary<string, string> relTargets = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
		using (Stream relsStream = relsEntry.Open())
		{
			XDocument rels = XDocument.Load(relsStream);
			XNamespace relNs = XNamespace.Get("http://schemas.openxmlformats.org/package/2006/relationships");
			foreach (XElement rel in rels.Root.Elements(relNs + "Relationship"))
			{
				string id = ((rel.Attribute("Id") == null) ? string.Empty : rel.Attribute("Id").Value);
				string target = ((rel.Attribute("Target") == null) ? string.Empty : rel.Attribute("Target").Value);
				if (!string.IsNullOrWhiteSpace(id) && !string.IsNullOrWhiteSpace(target))
				{
					relTargets[id] = NormalizeWorkbookTarget(target);
				}
			}
		}
		using (Stream workbookStream = workbookEntry.Open())
		{
			XDocument xDocument = XDocument.Load(workbookStream);
			XNamespace ns = XNamespace.Get("http://schemas.openxmlformats.org/spreadsheetml/2006/main");
			XNamespace relNs2 = XNamespace.Get("http://schemas.openxmlformats.org/officeDocument/2006/relationships");
			List<XElement> sheets = xDocument.Descendants(ns + "sheet").ToList();
			XElement selected = sheets.FirstOrDefault([SpecialName] (XElement sheet) => string.IsNullOrWhiteSpace(CS_0024_003C_003E8__locals3._0024VB_0024Local_requestedSheetName) || string.Equals((sheet.Attribute("name") == null) ? string.Empty : sheet.Attribute("name").Value, CS_0024_003C_003E8__locals3._0024VB_0024Local_requestedSheetName, StringComparison.OrdinalIgnoreCase));
			if (selected == null && sheets.Count > 0)
			{
				selected = sheets[0];
			}
			if (selected == null)
			{
				return string.Empty;
			}
			string relId = ((selected.Attribute(relNs2 + "id") == null) ? string.Empty : selected.Attribute(relNs2 + "id").Value);
			if (relTargets.ContainsKey(relId))
			{
				return relTargets[relId];
			}
		}
		return "xl/worksheets/sheet1.xml";
	}

	private static string NormalizeWorkbookTarget(string target)
	{
		string value = (target ?? string.Empty).Trim().Replace('\\', '/');
		if (value.StartsWith("/", StringComparison.Ordinal))
		{
			value = value.TrimStart('/');
		}
		else if (!value.StartsWith("xl/", StringComparison.OrdinalIgnoreCase))
		{
			value = "xl/" + value;
		}
		return value;
	}

	private static string ReadCellValue(XElement cell, XNamespace ns, List<string> sharedStrings)
	{
		string typeValue = ((cell.Attribute("t") == null) ? string.Empty : cell.Attribute("t").Value);
		if (string.Equals(typeValue, "inlineStr", StringComparison.OrdinalIgnoreCase))
		{
			return string.Concat(from x in cell.Descendants(ns + "t")
				select x.Value);
		}
		string raw = ((cell.Element(ns + "v") == null) ? string.Empty : cell.Element(ns + "v").Value);
		if (string.Equals(typeValue, "s", StringComparison.OrdinalIgnoreCase) && int.TryParse(raw, NumberStyles.Integer, CultureInfo.InvariantCulture, out var index) && index >= 0 && index < sharedStrings.Count)
		{
			return sharedStrings[index];
		}
		return raw;
	}

	private static int ResolveCellColumnIndex(string cellRef)
	{
		string letters = new string((cellRef ?? string.Empty).TakeWhile([SpecialName] (char c) => char.IsLetter(c)).ToArray()).ToUpperInvariant();
		if (string.IsNullOrWhiteSpace(letters))
		{
			return 0;
		}
		int result = 0;
		string text = letters;
		foreach (char ch in text)
		{
			result = checked(result * 26 + (ch - 65 + 1));
		}
		return result;
	}

	private static List<FamilyBrowserPermissionExcelRow> MapRows(List<string> header, List<List<string>> dataRows, int firstDataRowNumber)
	{
		Dictionary<int, string> map = new Dictionary<int, string>();
		checked
		{
			int num = header.Count - 1;
			for (int i = 0; i <= num; i++)
			{
				string normalized = NormalizeHeader(header[i]);
				if (!string.IsNullOrWhiteSpace(normalized))
				{
					map[i] = normalized;
				}
			}
			List<FamilyBrowserPermissionExcelRow> result = new List<FamilyBrowserPermissionExcelRow>();
			int num2 = dataRows.Count - 1;
			for (int j = 0; j <= num2; j++)
			{
				List<string> values = dataRows[j];
				if (values == null || values.All([SpecialName] (string x) => string.IsNullOrWhiteSpace(x)))
				{
					continue;
				}
				FamilyBrowserPermissionExcelRow row = new FamilyBrowserPermissionExcelRow
				{
					RowNumber = firstDataRowNumber + j
				};
				int num3 = values.Count - 1;
				for (int col = 0; col <= num3; col++)
				{
					if (map.ContainsKey(col))
					{
						ApplyCell(row, map[col], values[col]);
					}
				}
				if (!string.IsNullOrWhiteSpace(row.UserOrGroup) || !string.IsNullOrWhiteSpace(row.ApplyFolder) || !string.IsNullOrWhiteSpace(row.RvtFileName) || !string.IsNullOrWhiteSpace(row.ProjectName) || !string.IsNullOrWhiteSpace(row.CentralPathContains) || !string.IsNullOrWhiteSpace(row.MatchValue))
				{
					result.Add(row);
				}
			}
			return result;
		}
	}

	private static void ApplyCell(FamilyBrowserPermissionExcelRow row, string header, string value)
	{
		string text = (value ?? string.Empty).Trim();
		switch (header)
		{
		case "enabled":
		case "use":
		case "사용":
		case "사용여부":
			row.Enabled = text;
			return;
		case "applyfolder":
		case "folder":
		case "rootfolder":
		case "targetfolder":
		case "projectfolder":
		case "appliedfolder":
		case "적용폴더":
		case "대상폴더":
		case "프로젝트폴더":
			row.ApplyFolder = text;
			return;
		case "rvtfilename":
		case "rvtfile":
		case "filename":
		case "modelname":
		case "projectfile":
		case "rvt파일명":
		case "rvt파일":
		case "파일명":
		case "모델명":
		case "프로젝트파일":
			row.RvtFileName = text;
			return;
		case "projectkey":
		case "projectcode":
		case "프로젝트키":
		case "프로젝트코드":
			row.ProjectKey = text;
			return;
		case "projectname":
		case "projecttitle":
		case "프로젝트명":
		case "프로젝트":
			row.ProjectName = text;
			return;
		case "discipline":
		case "field":
		case "trade":
		case "공종":
		case "분야":
			row.Discipline = text;
			return;
		case "matchmode":
		case "매칭방식":
			row.MatchMode = text;
			return;
		case "matchvalue":
		case "매칭값":
			row.MatchValue = text;
			return;
		case "centralpath":
		case "central":
		case "센트럴경로":
			row.CentralPath = text;
			return;
		case "centralpathcontains":
		case "centralcontains":
		case "센트럴포함":
		case "경로포함":
			row.CentralPathContains = text;
			return;
		case "modelpath":
		case "rvtpath":
		case "파일경로":
		case "모델경로":
			row.ModelPath = text;
			return;
		case "modelpathcontains":
		case "rvtpathcontains":
		case "모델경로포함":
		case "파일경로포함":
			row.ModelPathContains = text;
			return;
		case "projecttitlecontains":
		case "프로젝트명포함":
			row.ProjectTitleContains = text;
			return;
		case "user":
		case "userorgroup":
		case "windowsuser":
		case "account":
		case "사용자":
		case "계정":
		case "사용자그룹":
			row.UserOrGroup = text;
			return;
		case "role":
		case "역할":
			row.Role = text;
			return;
		}
		string permission = ResolvePermissionNameFromHeader(header);
		if (!string.IsNullOrWhiteSpace(permission))
		{
			row.Permissions[permission] = text;
		}
	}

	private static string ResolvePermissionNameFromHeader(string header)
	{
		switch (NormalizeHeader(header))
		{
		case "editfamily":
		case "editfamilies":
		case "familyedit":
		case "loadeditfamily":
		case "loadeditfamilies":
		case "familyloadedit":
		case "nativefamilyload":
		case "nativefamilyedit":
		case "nativefamilyloadedit":
		case "패밀리로드편집":
		case "패밀리로드에디트":
		case "패밀리편집":
		case "패밀리에디트":
		case "패밀리로드":
			return "EditFamilies";
		case "adddeletetype":
		case "adddeletetypes":
		case "typeadddelete":
		case "nativeadddeletetype":
		case "nativeadddeletetypes":
		case "systemtypeadddelete":
		case "systemtypeadd":
		case "systemtypecreate":
		case "타입추가삭제":
		case "타입추가":
		case "타입삭제":
		case "시스템타입추가삭제":
		case "시스템타입추가":
			return "AddDeleteTypes";
		default:
			return string.Empty;
		}
	}

	private static string ResolvePermissionToken(FamilyBrowserPermissionExcelRow row, string permission)
	{
		if (!IsNativeGuardPermission(permission))
		{
			return string.Empty;
		}
		if (row == null || row.Permissions == null)
		{
			return string.Empty;
		}
		if (row.Permissions.ContainsKey(permission))
		{
			return row.Permissions[permission];
		}
		return string.Empty;
	}

	public static bool IsNativeGuardPermission(string permission)
	{
		if (!string.Equals(permission, "EditFamilies", StringComparison.OrdinalIgnoreCase))
		{
			return string.Equals(permission, "AddDeleteTypes", StringComparison.OrdinalIgnoreCase);
		}
		return true;
	}

	private static bool MatchesProject(FamilyBrowserPermissionExcelRow row, FamilyBrowserProjectPolicyContext context)
	{
		if (row == null)
		{
			return false;
		}
		if (context == null)
		{
			context = new FamilyBrowserProjectPolicyContext();
		}
		if (!string.IsNullOrWhiteSpace(row.Discipline) && !ContainsText(context.StandardTarget, row.Discipline))
		{
			return false;
		}
		if ((!string.IsNullOrWhiteSpace(row.ApplyFolder) || !string.IsNullOrWhiteSpace(row.RvtFileName)) && !MatchesApplyFolderRvtFile(row, context))
		{
			return false;
		}
		if ((!string.IsNullOrWhiteSpace(row.MatchMode) || !string.IsNullOrWhiteSpace(row.MatchValue)) && !MatchesByMode(row.MatchMode, row.MatchValue, context))
		{
			return false;
		}
		if (!string.IsNullOrWhiteSpace(row.CentralPath) && !SameText(context.CentralPath, row.CentralPath))
		{
			return false;
		}
		if (!string.IsNullOrWhiteSpace(row.CentralPathContains) && !ContainsText(context.CentralPath, row.CentralPathContains))
		{
			return false;
		}
		if (!string.IsNullOrWhiteSpace(row.ModelPath) && !SameText(context.ModelPath, row.ModelPath))
		{
			return false;
		}
		if (!string.IsNullOrWhiteSpace(row.ModelPathContains) && !ContainsText(context.ModelPath, row.ModelPathContains))
		{
			return false;
		}
		if (!string.IsNullOrWhiteSpace(row.ProjectTitleContains) && !ContainsText(context.ProjectTitle, row.ProjectTitleContains))
		{
			return false;
		}
		if (!string.IsNullOrWhiteSpace(row.ProjectName) && !ContainsText(context.ProjectTitle, row.ProjectName))
		{
			return false;
		}
		return true;
	}

	private static bool MatchesApplyFolderRvtFile(FamilyBrowserPermissionExcelRow row, FamilyBrowserProjectPolicyContext context)
	{
		if (row == null || context == null)
		{
			return false;
		}
		if (string.IsNullOrWhiteSpace(row.RvtFileName))
		{
			return false;
		}
		List<string> pathCandidates = BuildProjectPathCandidates(context);
		string rvtFileName = row.RvtFileName;
		if (!pathCandidates.Any([SpecialName] (string path) => RvtFileNameMatches(path, rvtFileName)))
		{
			return false;
		}
		if (string.IsNullOrWhiteSpace(row.ApplyFolder))
		{
			return true;
		}
		return pathCandidates.Any([SpecialName] (string path) => IsPathUnderFolder(path, row.ApplyFolder));
	}

	private static List<string> BuildProjectPathCandidates(FamilyBrowserProjectPolicyContext context)
	{
		List<string> values = new List<string>();
		AddPathCandidate(values, (context == null) ? string.Empty : context.CentralPath);
		AddPathCandidate(values, (context == null) ? string.Empty : context.ModelPath);
		AddPathCandidate(values, (context == null) ? string.Empty : context.ProjectTitle);
		if (context != null && !string.IsNullOrWhiteSpace(context.ProjectTitle))
		{
			AddPathCandidate(values, context.ProjectTitle + ".rvt");
		}
		return values;
	}

	private static void AddPathCandidate(List<string> values, string value)
	{
		string text = (value ?? string.Empty).Trim();
		if (!string.IsNullOrWhiteSpace(text) && !values.Any([SpecialName] (string x) => SameText(x, text)))
		{
			values.Add(text);
		}
	}

	private static bool RvtFileNameMatches(string candidatePath, string expectedFileName)
	{
		string candidateName = SafeFileName(candidatePath);
		string expectedName = SafeFileName(expectedFileName);
		if (string.IsNullOrWhiteSpace(candidateName) || string.IsNullOrWhiteSpace(expectedName))
		{
			return false;
		}
		if (SameText(candidateName, expectedName))
		{
			return true;
		}
		return SameText(Path.GetFileNameWithoutExtension(candidateName), Path.GetFileNameWithoutExtension(expectedName));
	}

	private static string SafeFileName(string value)
	{
		string text = (value ?? string.Empty).Trim().TrimEnd('\\', '/');
		if (string.IsNullOrWhiteSpace(text))
		{
			return string.Empty;
		}
		try
		{
			string name = Path.GetFileName(text);
			if (!string.IsNullOrWhiteSpace(name))
			{
				return name;
			}
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ProjectData.ClearProjectError();
		}
		int slashIndex = Math.Max(text.LastIndexOf('\\'), text.LastIndexOf('/'));
		checked
		{
			if (slashIndex >= 0 && slashIndex < text.Length - 1)
			{
				return text.Substring(slashIndex + 1);
			}
			return text;
		}
	}

	private static bool IsPathUnderFolder(string candidatePath, string applyFolder)
	{
		string pathText = Environment.ExpandEnvironmentVariables((candidatePath ?? string.Empty).Trim());
		string folderText = Environment.ExpandEnvironmentVariables((applyFolder ?? string.Empty).Trim());
		bool IsPathUnderFolder;
		if (string.IsNullOrWhiteSpace(pathText) || string.IsNullOrWhiteSpace(folderText))
		{
			IsPathUnderFolder = false;
		}
		else
		{
			try
			{
				string fullPath = Path.GetFullPath(pathText).TrimEnd('\\', '/');
				string fullFolder = Path.GetFullPath(folderText).TrimEnd('\\', '/');
				IsPathUnderFolder = SameText(fullPath, fullFolder) || fullPath.StartsWith(fullFolder + Conversions.ToString(Path.DirectorySeparatorChar), StringComparison.OrdinalIgnoreCase) || fullPath.StartsWith(fullFolder + Conversions.ToString(Path.AltDirectorySeparatorChar), StringComparison.OrdinalIgnoreCase);
			}
			catch (Exception projectError)
			{
				ProjectData.SetProjectError(projectError);
				string normalizedPath = pathText.Replace('/', '\\').TrimEnd('\\');
				string normalizedFolder = folderText.Replace('/', '\\').TrimEnd('\\');
				IsPathUnderFolder = SameText(normalizedPath, normalizedFolder) || normalizedPath.StartsWith(normalizedFolder + "\\", StringComparison.OrdinalIgnoreCase);
				ProjectData.ClearProjectError();
			}
		}
		return IsPathUnderFolder;
	}

	private static bool MatchesByMode(string matchMode, string matchValue, FamilyBrowserProjectPolicyContext context)
	{
		if (string.IsNullOrWhiteSpace(matchValue))
		{
			return true;
		}
		switch (FamilyBrowserPolicyKey.Normalize(matchMode))
		{
		case "any":
			return true;
		case "exactcentralpath":
		case "exact-central-path":
			return SameText(context.CentralPath, matchValue);
		case "exactmodelpath":
		case "exact-model-path":
			return SameText(context.ModelPath, matchValue);
		case "exactcentralormodelpath":
		case "exact-central-or-model-path":
		case "centralormodelpath":
		case "central-or-model-path":
		case "exactpath":
		case "exact-path":
			return SameText(context.CentralPath, matchValue) || SameText(context.ModelPath, matchValue);
		case "modelpathcontains":
		case "model-path-contains":
			return ContainsText(context.ModelPath, matchValue);
		case "pathcontains":
		case "path-contains":
		case "centralormodelpathcontains":
		case "central-or-model-path-contains":
			return ContainsText(context.CentralPath, matchValue) || ContainsText(context.ModelPath, matchValue);
		case "projecttitlecontains":
		case "project-title-contains":
		case "titlecontains":
			return ContainsText(context.ProjectTitle, matchValue);
		case "standardtarget":
		case "discipline":
			return ContainsText(context.StandardTarget, matchValue);
		default:
			return ContainsText(context.CentralPath, matchValue);
		}
	}

	private static bool MatchesUser(FamilyBrowserPermissionExcelRow row, string currentUser)
	{
		if (string.IsNullOrWhiteSpace(row.UserOrGroup))
		{
			return true;
		}
		HashSet<string> currentCandidates = BuildCurrentUserCandidates(currentUser);
		string[] array = row.UserOrGroup.Split(new char[5] { ',', ';', '\r', '\n', '\t' }, StringSplitOptions.RemoveEmptyEntries);
		for (int i = 0; i < array.Length; i = checked(i + 1))
		{
			string normalized = NormalizeUser(array[i]);
			if (string.Equals(normalized, "*", StringComparison.OrdinalIgnoreCase) || string.Equals(normalized, "all", StringComparison.OrdinalIgnoreCase))
			{
				return true;
			}
			if (currentCandidates.Contains(normalized))
			{
				return true;
			}
		}
		return false;
	}

	private static HashSet<string> BuildCurrentUserCandidates(string currentUser)
	{
		HashSet<string> values = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
		AddCandidate(values, currentUser);
		AddCandidate(values, Environment.UserName);
		AddCandidate(values, Environment.UserDomainName + "\\" + Environment.UserName);
		if (!string.IsNullOrWhiteSpace(currentUser) && currentUser.Contains("\\"))
		{
			AddCandidate(values, currentUser.Substring(checked(currentUser.LastIndexOf('\\') + 1)));
		}
		return values;
	}

	private static void AddCandidate(HashSet<string> values, string value)
	{
		string normalized = NormalizeUser(value);
		if (!string.IsNullOrWhiteSpace(normalized))
		{
			values.Add(normalized);
		}
	}

	private static string NormalizeUser(string value)
	{
		return (value ?? string.Empty).Trim().Replace('/', '\\').ToLowerInvariant();
	}

	private static bool IsRowEnabled(FamilyBrowserPermissionExcelRow row)
	{
		if (row == null || string.IsNullOrWhiteSpace(row.Enabled))
		{
			return true;
		}
		bool parsed = default(bool);
		if (TryParseYesNo(row.Enabled, ref parsed))
		{
			return parsed;
		}
		return true;
	}

	private static bool TryParseYesNo(string value, ref bool result)
	{
		switch (NormalizeHeader(value))
		{
		case "o":
		case "ok":
		case "yes":
		case "y":
		case "true":
		case "1":
		case "allow":
		case "allowed":
			result = true;
			return true;
		case "x":
		case "no":
		case "n":
		case "false":
		case "0":
		case "deny":
		case "denied":
		case "block":
		case "blocked":
			result = false;
			return true;
		default:
			return false;
		}
	}

	private static string NormalizeRole(string role)
	{
		switch (FamilyBrowserPolicyKey.Normalize(role))
		{
		case "admin":
		case "administrator":
		case "관리자":
			return "Admin";
		case "approver":
		case "requestapprover":
		case "request-approver":
		case "승인자":
			return "Approver";
		case "readonly":
		case "read-only":
		case "viewer":
		case "읽기전용":
			return "ReadOnly";
		case "modeler":
		case "모델러":
			return "Modeler";
		default:
			return string.Empty;
		}
	}

	private static List<string> SplitCsvLine(string line)
	{
		List<string> result = new List<string>();
		StringBuilder builder = new StringBuilder();
		bool inQuotes = false;
		checked
		{
			for (int i = 0; i < (line ?? string.Empty).Length; i++)
			{
				char ch = line[i];
				switch (ch)
				{
				case '"':
					if (inQuotes && i + 1 < line.Length && line[i + 1] == '"')
					{
						builder.Append('"');
						i++;
					}
					else
					{
						inQuotes = !inQuotes;
					}
					continue;
				case ',':
					if (!inQuotes)
					{
						result.Add(builder.ToString());
						builder.Clear();
						continue;
					}
					break;
				}
				builder.Append(ch);
			}
			result.Add(builder.ToString());
			return result;
		}
	}

	private static string ResolveExcelPath(FamilyBrowserStandardPolicy policy)
	{
		FamilyBrowserPermissionExcelSettings settings = policy?.PermissionExcel;
		if (settings == null)
		{
			return string.Empty;
		}
		return Environment.ExpandEnvironmentVariables((settings.Path ?? string.Empty).Trim());
	}

	private static string NormalizeHeader(string value)
	{
		return (value ?? string.Empty).Trim().Replace(" ", string.Empty).Replace("_", string.Empty)
			.Replace("-", string.Empty)
			.Replace("/", string.Empty)
			.Replace("\\", string.Empty)
			.ToLowerInvariant();
	}

	private static bool SameText(string leftValue, string rightValue)
	{
		return string.Equals((leftValue ?? string.Empty).Trim(), (rightValue ?? string.Empty).Trim(), StringComparison.OrdinalIgnoreCase);
	}

	private static bool ContainsText(string value, string token)
	{
		if (string.IsNullOrWhiteSpace(token))
		{
			return true;
		}
		return (value ?? string.Empty).IndexOf(token.Trim(), StringComparison.OrdinalIgnoreCase) >= 0;
	}
}
