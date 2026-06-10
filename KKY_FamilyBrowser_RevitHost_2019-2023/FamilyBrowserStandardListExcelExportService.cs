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

public sealed class FamilyBrowserStandardListExcelExportService
{
	private sealed class SignatureAuditRecord
	{
		public string CategoryName { get; set; }

		public string FamilyName { get; set; }

		public string Fingerprint { get; set; }

		public string ErrorMessage { get; set; }

		public string Path { get; set; }

		public long LastWriteUtcTicks { get; set; }

		public SignatureAuditRecord()
		{
			CategoryName = string.Empty;
			FamilyName = string.Empty;
			Fingerprint = string.Empty;
			ErrorMessage = string.Empty;
			Path = string.Empty;
		}
	}

	[CompilerGenerated]
	internal sealed class _Closure_0024__13_002D0
	{
		public string _0024VB_0024Local_disciplineKey;

		public string _0024VB_0024Local_disciplineLabel;

		public bool _0024VB_0024Local_hasDisciplineSpecificRows;

		public Func<FamilyBrowserStandardListEntry, bool> _0024I1;

		public _Closure_0024__13_002D0(_Closure_0024__13_002D0 arg0)
		{
			if (arg0 != null)
			{
				_0024VB_0024Local_disciplineKey = arg0._0024VB_0024Local_disciplineKey;
				_0024VB_0024Local_disciplineLabel = arg0._0024VB_0024Local_disciplineLabel;
				_0024VB_0024Local_hasDisciplineSpecificRows = arg0._0024VB_0024Local_hasDisciplineSpecificRows;
			}
		}

		[SpecialName]
		internal bool _Lambda_0024__0(FamilyBrowserStandardListEntry entry)
		{
			return StandardListDisciplineMatches((entry == null) ? string.Empty : entry.Discipline, _0024VB_0024Local_disciplineKey, _0024VB_0024Local_disciplineLabel);
		}

		[SpecialName]
		internal bool _Lambda_0024__1(FamilyBrowserStandardListEntry x)
		{
			if (x != null)
			{
				return StandardListEntryAppliesToDiscipline(x, _0024VB_0024Local_disciplineKey, _0024VB_0024Local_disciplineLabel, _0024VB_0024Local_hasDisciplineSpecificRows);
			}
			return false;
		}
	}

	[CompilerGenerated]
	internal sealed class _Closure_0024__34_002D0
	{
		public string _0024VB_0024Local_disciplineKey;

		public string _0024VB_0024Local_disciplineLabel;

		public _Closure_0024__34_002D0(_Closure_0024__34_002D0 arg0)
		{
			if (arg0 != null)
			{
				_0024VB_0024Local_disciplineKey = arg0._0024VB_0024Local_disciplineKey;
				_0024VB_0024Local_disciplineLabel = arg0._0024VB_0024Local_disciplineLabel;
			}
		}

		[SpecialName]
		internal bool _Lambda_0024__0(FamilyBrowserStandardListEntry entry)
		{
			return StandardListDisciplineMatches((entry == null) ? string.Empty : entry.Discipline, _0024VB_0024Local_disciplineKey, _0024VB_0024Local_disciplineLabel);
		}
	}

	[CompilerGenerated]
	internal sealed class _Closure_0024__35_002D0
	{
		public string _0024VB_0024Local_disciplineKey;

		public string _0024VB_0024Local_disciplineLabel;

		public _Closure_0024__35_002D0(_Closure_0024__35_002D0 arg0)
		{
			if (arg0 != null)
			{
				_0024VB_0024Local_disciplineKey = arg0._0024VB_0024Local_disciplineKey;
				_0024VB_0024Local_disciplineLabel = arg0._0024VB_0024Local_disciplineLabel;
			}
		}

		[SpecialName]
		internal bool _Lambda_0024__0(FamilyBrowserStandardListEntry entry)
		{
			return StandardListDisciplineMatches((entry == null) ? string.Empty : entry.Discipline, _0024VB_0024Local_disciplineKey, _0024VB_0024Local_disciplineLabel);
		}
	}

	public const string DefaultSheetName = "StandardList";

	public const string FingerprintAuditSheetName = "FingerprintAudit";

	private static readonly string[] StandardListHeaders = new string[5] { "분야", "Category", "Family", "Type", "비고" };

	private FamilyBrowserStandardListExcelExportService()
	{
	}

	public static FamilyBrowserStandardListExcelExportResult SaveTemplate(string outputPath, string defaultDiscipline)
	{
		string discipline = (string.IsNullOrWhiteSpace(defaultDiscipline) ? "PI" : defaultDiscipline);
		List<List<string>> rows = new List<List<string>>
		{
			new List<string>
			{
				discipline,
				"Pipe Fittings",
				"Short Elbow, STS",
				string.Empty,
				"Loadable family example. Leave Type blank; loadable family filtering is family-level."
			},
			new List<string>
			{
				discipline,
				"Pipe Accessories",
				"Control Valve_Ball_Flange, ST'L+PTFE, NPS_3, Stub End_PE, ACV",
				string.Empty,
				"Loadable family example."
			},
			new List<string> { discipline, "Pipes", "PipeType", "Default", "System type example. Type is matched for PipeType/DuctType/etc." },
			new List<string> { discipline, "Ducts", "DuctType", "Default", "System type example." }
		};
		WriteWorkbook(outputPath, "StandardList", StandardListHeaders, rows);
		return new FamilyBrowserStandardListExcelExportResult
		{
			OutputPath = outputPath,
			RowCount = rows.Count,
			SheetName = "StandardList"
		};
	}

	public static FamilyBrowserStandardListExcelExportResult SaveFromSnapshot(string outputPath, string defaultDiscipline, StandardLibrarySnapshot snapshot, FamilyBrowserStandardListCatalog catalog = null, string disciplineKey = "", string disciplineLabel = "")
	{
		if (snapshot == null)
		{
			throw new ArgumentNullException("snapshot");
		}
		List<List<string>> rows = BuildRowsFromSnapshot(defaultDiscipline, snapshot, catalog, disciplineKey, disciplineLabel);
		WriteWorkbook(outputPath, "StandardList", StandardListHeaders, rows);
		return new FamilyBrowserStandardListExcelExportResult
		{
			OutputPath = outputPath,
			RowCount = rows.Count,
			SheetName = "StandardList"
		};
	}

	public static FamilyBrowserStandardFingerprintAuditExportResult SaveFingerprintAudit(string outputPath, string defaultDiscipline, FamilyBrowserStandardListCatalog catalog, StandardLibrarySnapshot snapshot, string disciplineKey, string disciplineLabel, bool korean = false)
	{
		if (catalog == null)
		{
			throw new ArgumentNullException("catalog");
		}
		if (snapshot == null)
		{
			throw new ArgumentNullException("snapshot");
		}
		FamilyBrowserStandardFingerprintAuditExportResult result = new FamilyBrowserStandardFingerprintAuditExportResult
		{
			OutputPath = outputPath,
			SheetName = "FingerprintAudit"
		};
		List<List<string>> rows = BuildFingerprintAuditRows(defaultDiscipline, catalog, snapshot, disciplineKey, disciplineLabel, korean, result);
		WriteWorkbook(outputPath, "FingerprintAudit", BuildFingerprintAuditHeadersForExport(korean), rows);
		result.RowCount = rows.Count;
		return result;
	}

	public static FamilyBrowserStandardListExcelExportResult SaveRows(string outputPath, string sheetName, IList<string> headers, IList<List<string>> rows)
	{
		IList<string> safeHeaders = headers ?? new List<string>();
		IList<List<string>> safeRows = rows ?? new List<List<string>>();
		string safeSheetName = (string.IsNullOrWhiteSpace(sheetName) ? "StandardList" : sheetName.Trim());
		WriteWorkbook(outputPath, safeSheetName, safeHeaders, safeRows);
		return new FamilyBrowserStandardListExcelExportResult
		{
			OutputPath = outputPath,
			RowCount = safeRows.Count,
			SheetName = safeSheetName
		};
	}

	private static List<List<string>> BuildRowsFromSnapshot(string defaultDiscipline, StandardLibrarySnapshot snapshot, FamilyBrowserStandardListCatalog catalog, string disciplineKey, string disciplineLabel)
	{
		string discipline = (string.IsNullOrWhiteSpace(defaultDiscipline) ? "PI" : defaultDiscipline.Trim());
		List<List<string>> rows = new List<List<string>>();
		HashSet<string> nestedFamilyNames = BuildNestedLoadableNameSet(snapshot);
		HashSet<string> loadableKeys = new HashSet<string>(StringComparer.Ordinal);
		foreach (StandardLoadableFamilySnapshotItem item in (snapshot.LoadableFamilies ?? new List<StandardLoadableFamilySnapshotItem>()).Where([SpecialName] (StandardLoadableFamilySnapshotItem x) => x != null).OrderBy([SpecialName] (StandardLoadableFamilySnapshotItem x) => Normalize(x.CategoryName), StringComparer.Ordinal).ThenBy([SpecialName] (StandardLoadableFamilySnapshotItem x) => Normalize(x.FamilyName), StringComparer.Ordinal))
		{
			if (!ShouldSkipLoadableFamily(item, nestedFamilyNames) && !StandardListBaselineExcludesLoadable(catalog, disciplineKey, disciplineLabel, item.CategoryName, item.FamilyName))
			{
				string key = Normalize(item.CategoryName) + "|" + Normalize(item.FamilyName);
				if (!loadableKeys.Contains(key))
				{
					loadableKeys.Add(key);
					rows.Add(new List<string>
					{
						discipline,
						item.CategoryName ?? string.Empty,
						item.FamilyName ?? string.Empty,
						string.Empty,
						BuildLoadableNote(item)
					});
				}
			}
		}
		HashSet<string> systemKeys = new HashSet<string>(StringComparer.Ordinal);
		foreach (StandardSystemTypeSnapshotItem item2 in (snapshot.SystemTypes ?? new List<StandardSystemTypeSnapshotItem>()).Where([SpecialName] (StandardSystemTypeSnapshotItem x) => x != null).OrderBy([SpecialName] (StandardSystemTypeSnapshotItem x) => Normalize(x.CategoryName), StringComparer.Ordinal).ThenBy([SpecialName] (StandardSystemTypeSnapshotItem x) => Normalize(x.TypeClassName), StringComparer.Ordinal)
			.ThenBy([SpecialName] (StandardSystemTypeSnapshotItem x) => Normalize(x.TypeName), StringComparer.Ordinal))
		{
			string key2 = Normalize(item2.CategoryName) + "|" + Normalize(item2.TypeClassName) + "|" + Normalize(item2.TypeName);
			if (!systemKeys.Contains(key2) && !StandardListBaselineExcludesSystem(catalog, disciplineKey, disciplineLabel, item2.CategoryName, item2.TypeClassName, item2.TypeName))
			{
				systemKeys.Add(key2);
				rows.Add(new List<string>
				{
					discipline,
					item2.CategoryName ?? string.Empty,
					item2.TypeClassName ?? string.Empty,
					item2.TypeName ?? string.Empty,
					"System type from registered standard RVT."
				});
			}
		}
		return rows;
	}

	private static List<string> BuildFingerprintAuditHeaders(bool korean)
	{
		if (korean)
		{
			return new List<string> { "분야", "Category", "Family", "Type", "비고", "Fingerprint 상태", "Fingerprint", "Signature 경로" };
		}
		return new List<string> { "Discipline", "Category", "Family", "Type", "Notes", "Fingerprint Status", "Fingerprint", "Signature Path" };
	}

	private static List<string> BuildFingerprintAuditHeadersForExport(bool korean)
	{
		if (korean)
		{
			return new List<string> { "분야", "Category", "Family", "Type", "비고", "Fingerprint 상태", "Fingerprint", "Signature 경로" };
		}
		return BuildFingerprintAuditHeaders(korean: false);
	}

	private static List<List<string>> BuildFingerprintAuditRows(string defaultDiscipline, FamilyBrowserStandardListCatalog catalog, StandardLibrarySnapshot snapshot, string disciplineKey, string disciplineLabel, bool korean, FamilyBrowserStandardFingerprintAuditExportResult result)
	{
		_Closure_0024__13_002D0 arg = default(_Closure_0024__13_002D0);
		_Closure_0024__13_002D0 CS_0024_003C_003E8__locals10 = new _Closure_0024__13_002D0(arg);
		CS_0024_003C_003E8__locals10._0024VB_0024Local_disciplineKey = disciplineKey;
		CS_0024_003C_003E8__locals10._0024VB_0024Local_disciplineLabel = disciplineLabel;
		List<List<string>> rows = new List<List<string>>();
		List<FamilyBrowserStandardListEntry> entries = catalog.Entries ?? new List<FamilyBrowserStandardListEntry>();
		Dictionary<string, List<StandardLoadableFamilySnapshotItem>> loadableIndex = BuildLoadableSnapshotIndex(snapshot);
		Dictionary<string, List<StandardSystemTypeSnapshotItem>> systemIndex = BuildSystemTypeSnapshotIndex(snapshot);
		Dictionary<string, List<SignatureAuditRecord>> signatureIndex = BuildSignatureAuditIndex(snapshot);
		CS_0024_003C_003E8__locals10._0024VB_0024Local_hasDisciplineSpecificRows = entries.Any([SpecialName] (FamilyBrowserStandardListEntry familyBrowserStandardListEntry) => StandardListDisciplineMatches((familyBrowserStandardListEntry == null) ? string.Empty : familyBrowserStandardListEntry.Discipline, CS_0024_003C_003E8__locals10._0024VB_0024Local_disciplineKey, CS_0024_003C_003E8__locals10._0024VB_0024Local_disciplineLabel));
		string defaultTarget = FirstNonEmpty(defaultDiscipline, CS_0024_003C_003E8__locals10._0024VB_0024Local_disciplineLabel, CS_0024_003C_003E8__locals10._0024VB_0024Local_disciplineKey, "Standard");
		checked
		{
			foreach (FamilyBrowserStandardListEntry entry in entries.Where([SpecialName] (FamilyBrowserStandardListEntry x) => x != null && StandardListEntryAppliesToDiscipline(x, CS_0024_003C_003E8__locals10._0024VB_0024Local_disciplineKey, CS_0024_003C_003E8__locals10._0024VB_0024Local_disciplineLabel, CS_0024_003C_003E8__locals10._0024VB_0024Local_hasDisciplineSpecificRows)).OrderBy([SpecialName] (FamilyBrowserStandardListEntry x) => x.Category ?? string.Empty, StringComparer.OrdinalIgnoreCase).ThenBy([SpecialName] (FamilyBrowserStandardListEntry x) => x.Family ?? string.Empty, StringComparer.OrdinalIgnoreCase)
				.ThenBy([SpecialName] (FamilyBrowserStandardListEntry x) => x.TypeName ?? string.Empty, StringComparer.OrdinalIgnoreCase))
			{
				StandardLoadableFamilySnapshotItem matchedLoadable = FindMatchingLoadable(loadableIndex, entry);
				StandardSystemTypeSnapshotItem matchedSystem = FindMatchingSystemType(systemIndex, entry);
				string fingerprint = string.Empty;
				string signaturePath = string.Empty;
				string statusText;
				if (matchedLoadable != null)
				{
					fingerprint = matchedLoadable.ContentFingerprint ?? string.Empty;
					signaturePath = matchedLoadable.ContentSignatureDebugPath ?? string.Empty;
					SignatureAuditRecord signatureRecord = FindMatchingSignatureRecord(signatureIndex, matchedLoadable);
					if (string.IsNullOrWhiteSpace(signaturePath) && signatureRecord != null)
					{
						signaturePath = signatureRecord.Path ?? string.Empty;
					}
					if (string.IsNullOrWhiteSpace(fingerprint))
					{
						if (signatureRecord != null && !string.IsNullOrWhiteSpace(signatureRecord.Fingerprint))
						{
							fingerprint = signatureRecord.Fingerprint;
							result.RecoveredFingerprintCount++;
							statusText = (korean ? "스냅샷 JSON의 Fingerprint가 비어 있어 signature 파일에서 복구했습니다." : "Snapshot JSON fingerprint was empty; recovered it from the signature file.");
						}
						else
						{
							result.MissingFingerprintCount++;
							string failureReason = FirstNonEmpty(matchedLoadable.ContentFingerprintFailureReason ?? string.Empty, (signatureRecord == null) ? string.Empty : signatureRecord.ErrorMessage);
							statusText = (string.IsNullOrWhiteSpace(failureReason) ? (string.IsNullOrWhiteSpace(signaturePath) ? (korean ? "Fingerprint가 생성되지 않았고 signature 파일도 없습니다." : "Fingerprint was not created and no signature file was recorded.") : (korean ? "Signature 파일은 있으나 content-fingerprint 값이 비어 있습니다." : "Signature file exists, but its content-fingerprint value is empty.")) : (korean ? ("Fingerprint가 생성되지 않았습니다: " + failureReason) : ("Fingerprint was not created: " + failureReason)));
						}
					}
					else
					{
						statusText = (korean ? "Fingerprint 생성됨" : "Fingerprint created");
					}
				}
				else if (matchedSystem != null)
				{
					result.SystemTypeRowCount++;
					fingerprint = ProjectSnapshotFingerprintService.BuildSystemFingerprint(matchedSystem);
					statusText = (korean ? "시스템 타입 행입니다. 로더블 패밀리 Fingerprint 대상이 아닙니다." : "System type row; loadable family fingerprint is not applicable.");
				}
				else
				{
					result.MissingFromSnapshotCount++;
					statusText = (korean ? "표준 RVT 스냅샷에서 해당 항목을 찾지 못했습니다. Fingerprint가 리스트에 없습니다." : "Matching item was not found in the standard RVT snapshot. Fingerprint is not in the list.");
				}
				rows.Add(new List<string>
				{
					FirstNonEmpty(entry.Discipline, defaultTarget),
					entry.Category ?? string.Empty,
					entry.Family ?? string.Empty,
					entry.TypeName ?? string.Empty,
					AppendAuditNote(entry.Notes, statusText),
					statusText,
					fingerprint,
					signaturePath
				});
			}
			return rows;
		}
	}

	private static Dictionary<string, List<StandardLoadableFamilySnapshotItem>> BuildLoadableSnapshotIndex(StandardLibrarySnapshot snapshot)
	{
		Dictionary<string, List<StandardLoadableFamilySnapshotItem>> index = new Dictionary<string, List<StandardLoadableFamilySnapshotItem>>(StringComparer.Ordinal);
		if (snapshot == null || snapshot.LoadableFamilies == null)
		{
			return index;
		}
		foreach (StandardLoadableFamilySnapshotItem item in snapshot.LoadableFamilies)
		{
			if (item != null && !string.IsNullOrWhiteSpace(item.FamilyName))
			{
				string key = FoldToken(item.FamilyName);
				List<StandardLoadableFamilySnapshotItem> entries = null;
				if (!index.TryGetValue(key, out entries))
				{
					entries = (index[key] = new List<StandardLoadableFamilySnapshotItem>());
				}
				entries.Add(item);
			}
		}
		return index;
	}

	private static Dictionary<string, List<StandardSystemTypeSnapshotItem>> BuildSystemTypeSnapshotIndex(StandardLibrarySnapshot snapshot)
	{
		Dictionary<string, List<StandardSystemTypeSnapshotItem>> index = new Dictionary<string, List<StandardSystemTypeSnapshotItem>>(StringComparer.Ordinal);
		if (snapshot == null || snapshot.SystemTypes == null)
		{
			return index;
		}
		foreach (StandardSystemTypeSnapshotItem item in snapshot.SystemTypes)
		{
			if (item != null && !string.IsNullOrWhiteSpace(item.TypeName))
			{
				string key = FoldToken(item.TypeName);
				List<StandardSystemTypeSnapshotItem> entries = null;
				if (!index.TryGetValue(key, out entries))
				{
					entries = (index[key] = new List<StandardSystemTypeSnapshotItem>());
				}
				entries.Add(item);
			}
		}
		return index;
	}

	private static Dictionary<string, List<SignatureAuditRecord>> BuildSignatureAuditIndex(StandardLibrarySnapshot snapshot)
	{
		Dictionary<string, List<SignatureAuditRecord>> index = new Dictionary<string, List<SignatureAuditRecord>>(StringComparer.Ordinal);
		if (snapshot == null || snapshot.LoadableFamilies == null)
		{
			return index;
		}
		HashSet<string> loadableFolders = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
		foreach (StandardLoadableFamilySnapshotItem item in snapshot.LoadableFamilies)
		{
			if (item == null || string.IsNullOrWhiteSpace(item.ContentSignatureDebugPath))
			{
				continue;
			}
			AddSignatureAuditFile(index, item.ContentSignatureDebugPath);
			try
			{
				string parentFolder = Path.GetDirectoryName(Environment.ExpandEnvironmentVariables(item.ContentSignatureDebugPath));
				if (!string.IsNullOrWhiteSpace(parentFolder) && string.Equals(Path.GetFileName(parentFolder), "LoadableFamilies", StringComparison.OrdinalIgnoreCase))
				{
					loadableFolders.Add(parentFolder);
				}
			}
			catch (Exception projectError)
			{
				ProjectData.SetProjectError(projectError);
				ProjectData.ClearProjectError();
			}
		}
		foreach (string folder in loadableFolders)
		{
			try
			{
				if (Directory.Exists(folder))
				{
					string[] files = Directory.GetFiles(folder, "*.signature.txt", SearchOption.TopDirectoryOnly);
					foreach (string path in files)
					{
						AddSignatureAuditFile(index, path);
					}
				}
			}
			catch (Exception projectError2)
			{
				ProjectData.SetProjectError(projectError2);
				ProjectData.ClearProjectError();
			}
		}
		return index;
	}

	private static void AddSignatureAuditFile(Dictionary<string, List<SignatureAuditRecord>> index, string signaturePath)
	{
		if (index == null || string.IsNullOrWhiteSpace(signaturePath))
		{
			return;
		}
		SignatureAuditRecord record = ReadSignatureAuditRecord(signaturePath);
		if (record != null && !string.IsNullOrWhiteSpace(record.FamilyName))
		{
			AddSignatureAuditRecord(index, "f|" + FoldToken(record.FamilyName), record);
			if (!string.IsNullOrWhiteSpace(record.CategoryName))
			{
				AddSignatureAuditRecord(index, "cf|" + FoldToken(record.CategoryName) + "|" + FoldToken(record.FamilyName), record);
			}
		}
	}

	private static void AddSignatureAuditRecord(Dictionary<string, List<SignatureAuditRecord>> index, string key, SignatureAuditRecord record)
	{
		if (!string.IsNullOrWhiteSpace(key) && record != null)
		{
			List<SignatureAuditRecord> records = null;
			if (!index.TryGetValue(key, out records))
			{
				records = (index[key] = new List<SignatureAuditRecord>());
			}
			if (!records.Any([SpecialName] (SignatureAuditRecord x) => x != null && string.Equals(x.Path, record.Path, StringComparison.OrdinalIgnoreCase)))
			{
				records.Add(record);
			}
		}
	}

	private static SignatureAuditRecord ReadSignatureAuditRecord(string signaturePath)
	{
		SignatureAuditRecord ReadSignatureAuditRecord;
		try
		{
			string expandedPath = Environment.ExpandEnvironmentVariables((signaturePath ?? string.Empty).Trim());
			if (string.IsNullOrWhiteSpace(expandedPath) || !File.Exists(expandedPath))
			{
				ReadSignatureAuditRecord = null;
			}
			else
			{
				SignatureAuditRecord record = new SignatureAuditRecord
				{
					Path = expandedPath
				};
				try
				{
					record.LastWriteUtcTicks = new FileInfo(expandedPath).LastWriteTimeUtc.Ticks;
				}
				catch (Exception projectError)
				{
					ProjectData.SetProjectError(projectError);
					ProjectData.ClearProjectError();
				}
				foreach (string line in File.ReadLines(expandedPath, Encoding.UTF8))
				{
					if (string.IsNullOrWhiteSpace(line))
					{
						continue;
					}
					if (line.StartsWith("-----", StringComparison.Ordinal))
					{
						break;
					}
					int separator = line.IndexOf('=');
					if (separator > 0)
					{
						string text = line.Substring(0, separator).Trim();
						string value = line.Substring(checked(separator + 1)).Trim();
						switch (text.ToLowerInvariant())
						{
						case "category":
							record.CategoryName = value;
							break;
						case "family":
							record.FamilyName = value;
							break;
						case "content-fingerprint":
							record.Fingerprint = value;
							break;
						case "error-message":
							record.ErrorMessage = value;
							break;
						}
					}
				}
				ReadSignatureAuditRecord = record;
			}
		}
		catch (Exception projectError2)
		{
			ProjectData.SetProjectError(projectError2);
			ReadSignatureAuditRecord = null;
			ProjectData.ClearProjectError();
		}
		return ReadSignatureAuditRecord;
	}

	private static SignatureAuditRecord FindMatchingSignatureRecord(Dictionary<string, List<SignatureAuditRecord>> index, StandardLoadableFamilySnapshotItem item)
	{
		if (index == null || item == null || string.IsNullOrWhiteSpace(item.FamilyName))
		{
			return null;
		}
		string exactKey = "cf|" + FoldToken(item.CategoryName) + "|" + FoldToken(item.FamilyName);
		List<SignatureAuditRecord> records = null;
		if (index.TryGetValue(exactKey, out records) && records != null && records.Count > 0)
		{
			return SelectBestSignatureRecord(records);
		}
		string familyKey = "f|" + FoldToken(item.FamilyName);
		if (index.TryGetValue(familyKey, out records) && records != null && records.Count > 0)
		{
			return SelectBestSignatureRecord(records);
		}
		return null;
	}

	private static SignatureAuditRecord SelectBestSignatureRecord(IEnumerable<SignatureAuditRecord> records)
	{
		return (from x in records?.Where([SpecialName] (SignatureAuditRecord x) => x != null)
			orderby !string.IsNullOrWhiteSpace(x.Fingerprint) descending, x.LastWriteUtcTicks descending
			select x).FirstOrDefault();
	}

	private static StandardLoadableFamilySnapshotItem FindMatchingLoadable(Dictionary<string, List<StandardLoadableFamilySnapshotItem>> index, FamilyBrowserStandardListEntry entry)
	{
		if (index == null || entry == null || string.IsNullOrWhiteSpace(entry.Family))
		{
			return null;
		}
		List<StandardLoadableFamilySnapshotItem> candidates = null;
		if (!index.TryGetValue(FoldToken(entry.Family), out candidates) || candidates == null)
		{
			return null;
		}
		return candidates.FirstOrDefault([SpecialName] (StandardLoadableFamilySnapshotItem item) => item != null && CategoryMatches(entry.Category, item.CategoryName)) ?? candidates.FirstOrDefault([SpecialName] (StandardLoadableFamilySnapshotItem item) => item != null);
	}

	private static StandardSystemTypeSnapshotItem FindMatchingSystemType(Dictionary<string, List<StandardSystemTypeSnapshotItem>> index, FamilyBrowserStandardListEntry entry)
	{
		if (index == null || entry == null || string.IsNullOrWhiteSpace(entry.TypeName))
		{
			return null;
		}
		List<StandardSystemTypeSnapshotItem> candidates = null;
		if (!index.TryGetValue(FoldToken(entry.TypeName), out candidates) || candidates == null)
		{
			return null;
		}
		return candidates.FirstOrDefault([SpecialName] (StandardSystemTypeSnapshotItem item) => item != null && CategoryMatches(entry.Category, item.CategoryName) && SystemFamilyMatches(entry.Family, item.TypeClassName, item.CategoryName));
	}

	private static bool StandardListEntryAppliesToDiscipline(FamilyBrowserStandardListEntry entry, string disciplineKey, string disciplineLabel, bool hasDisciplineSpecificRows)
	{
		if (entry == null)
		{
			return false;
		}
		if (string.IsNullOrWhiteSpace(entry.Discipline) || !hasDisciplineSpecificRows)
		{
			return true;
		}
		return StandardListDisciplineMatches(entry.Discipline, disciplineKey, disciplineLabel);
	}

	private static bool StandardListDisciplineMatches(string entryDiscipline, string disciplineKey, string disciplineLabel)
	{
		if (string.IsNullOrWhiteSpace(entryDiscipline))
		{
			return true;
		}
		return SimilarToken(entryDiscipline, disciplineKey) || SimilarToken(entryDiscipline, disciplineLabel);
	}

	private static bool CategoryMatches(string entryCategory, string categoryName)
	{
		if (string.IsNullOrWhiteSpace(entryCategory))
		{
			return true;
		}
		return SimilarToken(entryCategory, categoryName);
	}

	private static bool SystemFamilyMatches(string entryFamily, string systemFamilyKind, string categoryName)
	{
		if (string.IsNullOrWhiteSpace(entryFamily))
		{
			return true;
		}
		return SimilarToken(entryFamily, systemFamilyKind) || SimilarToken(entryFamily, categoryName);
	}

	private static bool SimilarToken(string leftValue, string rightValue)
	{
		string left = FoldToken(leftValue);
		string right = FoldToken(rightValue);
		if (string.IsNullOrWhiteSpace(left) || string.IsNullOrWhiteSpace(right))
		{
			return false;
		}
		if (string.Equals(left, right, StringComparison.Ordinal))
		{
			return true;
		}
		checked
		{
			if (left.EndsWith("s", StringComparison.Ordinal) && string.Equals(left.Substring(0, left.Length - 1), right, StringComparison.Ordinal))
			{
				return true;
			}
			if (right.EndsWith("s", StringComparison.Ordinal) && string.Equals(right.Substring(0, right.Length - 1), left, StringComparison.Ordinal))
			{
				return true;
			}
			return left.Length >= 5 && right.Length >= 5 && (left.IndexOf(right, StringComparison.Ordinal) >= 0 || right.IndexOf(left, StringComparison.Ordinal) >= 0);
		}
	}

	private static string FoldToken(string value)
	{
		return (value ?? string.Empty).Trim().Replace(" ", string.Empty).Replace("_", string.Empty)
			.Replace("-", string.Empty)
			.Replace(".", string.Empty)
			.Replace("/", string.Empty)
			.Replace("\\", string.Empty)
			.ToLowerInvariant();
	}

	private static string AppendAuditNote(string originalNote, string auditNote)
	{
		if (string.IsNullOrWhiteSpace(originalNote))
		{
			return auditNote ?? string.Empty;
		}
		if (string.IsNullOrWhiteSpace(auditNote))
		{
			return originalNote;
		}
		return originalNote.Trim() + " / " + auditNote.Trim();
	}

	private static string FirstNonEmpty(params string[] values)
	{
		if (values == null)
		{
			return string.Empty;
		}
		foreach (string value in values)
		{
			if (!string.IsNullOrWhiteSpace(value))
			{
				return value.Trim();
			}
		}
		return string.Empty;
	}

	private static bool ShouldSkipLoadableFamily(StandardLoadableFamilySnapshotItem item, HashSet<string> nestedFamilyNames)
	{
		if (item == null)
		{
			return true;
		}
		if (FamilyBrowserFamilyClassificationService.IsTypeManagedFamilyLike(item.CategoryName, item.CategoryId, item.FamilyName))
		{
			return true;
		}
		return item.IsShared && IsModelLoadableFamily(item) && (item.IsNestedLoadableChild || (nestedFamilyNames?.Contains(Normalize(item.FamilyName)) ?? false));
	}

	private static string BuildLoadableNote(StandardLoadableFamilySnapshotItem item)
	{
		int typeCount = ((item != null) ? (item.TypeNames ?? new List<string>()).Count : 0);
		int nestedCount = ((item != null && item.NestedLoadableFamilies != null) ? item.NestedLoadableFamilies.Where([SpecialName] (StandardNestedLoadableFamilySnapshotItem x) => x != null && x.IsShared && IsModelNestedLoadableChild(x)).Count() : 0);
		string note = "Loadable family from registered standard RVT. Type count: " + typeCount.ToString(CultureInfo.InvariantCulture) + ".";
		if (nestedCount > 0)
		{
			note = note + " Composite parent; nested child families omitted: " + nestedCount.ToString(CultureInfo.InvariantCulture) + ".";
		}
		return note;
	}

	private static bool StandardListBaselineExcludesLoadable(FamilyBrowserStandardListCatalog catalog, string disciplineKey, string disciplineLabel, string categoryName, string familyName)
	{
		_Closure_0024__34_002D0 arg = default(_Closure_0024__34_002D0);
		_Closure_0024__34_002D0 CS_0024_003C_003E8__locals6 = new _Closure_0024__34_002D0(arg);
		CS_0024_003C_003E8__locals6._0024VB_0024Local_disciplineKey = disciplineKey;
		CS_0024_003C_003E8__locals6._0024VB_0024Local_disciplineLabel = disciplineLabel;
		if (catalog == null || catalog.BaselineExcludedLoadableFamilies == null || catalog.BaselineExcludedLoadableFamilies.Count == 0)
		{
			return false;
		}
		bool hasDisciplineSpecificRows = catalog.BaselineExcludedLoadableFamilies.Any([SpecialName] (FamilyBrowserStandardListEntry familyBrowserStandardListEntry) => StandardListDisciplineMatches((familyBrowserStandardListEntry == null) ? string.Empty : familyBrowserStandardListEntry.Discipline, CS_0024_003C_003E8__locals6._0024VB_0024Local_disciplineKey, CS_0024_003C_003E8__locals6._0024VB_0024Local_disciplineLabel));
		foreach (FamilyBrowserStandardListEntry entry in catalog.BaselineExcludedLoadableFamilies)
		{
			if (entry != null && StandardListEntryAppliesToDiscipline(entry, CS_0024_003C_003E8__locals6._0024VB_0024Local_disciplineKey, CS_0024_003C_003E8__locals6._0024VB_0024Local_disciplineLabel, hasDisciplineSpecificRows) && SameToken(entry.Family, familyName) && (string.IsNullOrWhiteSpace(entry.Category) || SimilarToken(entry.Category, categoryName)))
			{
				return true;
			}
		}
		return false;
	}

	private static bool StandardListBaselineExcludesSystem(FamilyBrowserStandardListCatalog catalog, string disciplineKey, string disciplineLabel, string categoryName, string systemFamilyKind, string typeName)
	{
		_Closure_0024__35_002D0 arg = default(_Closure_0024__35_002D0);
		_Closure_0024__35_002D0 CS_0024_003C_003E8__locals6 = new _Closure_0024__35_002D0(arg);
		CS_0024_003C_003E8__locals6._0024VB_0024Local_disciplineKey = disciplineKey;
		CS_0024_003C_003E8__locals6._0024VB_0024Local_disciplineLabel = disciplineLabel;
		if (catalog == null || catalog.BaselineExcludedSystemTypes == null || catalog.BaselineExcludedSystemTypes.Count == 0)
		{
			return false;
		}
		bool hasDisciplineSpecificRows = catalog.BaselineExcludedSystemTypes.Any([SpecialName] (FamilyBrowserStandardListEntry familyBrowserStandardListEntry) => StandardListDisciplineMatches((familyBrowserStandardListEntry == null) ? string.Empty : familyBrowserStandardListEntry.Discipline, CS_0024_003C_003E8__locals6._0024VB_0024Local_disciplineKey, CS_0024_003C_003E8__locals6._0024VB_0024Local_disciplineLabel));
		foreach (FamilyBrowserStandardListEntry entry in catalog.BaselineExcludedSystemTypes)
		{
			if (entry != null && StandardListEntryAppliesToDiscipline(entry, CS_0024_003C_003E8__locals6._0024VB_0024Local_disciplineKey, CS_0024_003C_003E8__locals6._0024VB_0024Local_disciplineLabel, hasDisciplineSpecificRows))
			{
				bool num = string.IsNullOrWhiteSpace(entry.Category) || SimilarToken(entry.Category, categoryName);
				bool familyMatches = string.IsNullOrWhiteSpace(entry.Family) || SimilarToken(entry.Family, systemFamilyKind) || SimilarToken(entry.Family, categoryName);
				if (num && familyMatches && SameToken(entry.TypeName, typeName))
				{
					return true;
				}
			}
		}
		return false;
	}

	private static bool SameToken(string leftValue, string rightValue)
	{
		string left = FoldToken(leftValue);
		string right = FoldToken(rightValue);
		if (left.Length > 0)
		{
			return string.Equals(left, right, StringComparison.Ordinal);
		}
		return false;
	}

	private static HashSet<string> BuildNestedLoadableNameSet(StandardLibrarySnapshot snapshot)
	{
		HashSet<string> result = new HashSet<string>(StringComparer.Ordinal);
		if (snapshot == null || snapshot.LoadableFamilies == null)
		{
			return result;
		}
		foreach (StandardLoadableFamilySnapshotItem parentItem in snapshot.LoadableFamilies)
		{
			if (parentItem == null)
			{
				continue;
			}
			if (parentItem.IsNestedLoadableChild && parentItem.IsShared && IsModelLoadableFamily(parentItem))
			{
				string nestedFamilyName = Normalize(parentItem.FamilyName);
				if (nestedFamilyName.Length > 0)
				{
					result.Add(nestedFamilyName);
				}
			}
			if (parentItem.NestedLoadableFamilies == null)
			{
				continue;
			}
			foreach (StandardNestedLoadableFamilySnapshotItem child in parentItem.NestedLoadableFamilies)
			{
				if (child != null && child.IsShared && IsModelNestedLoadableChild(child))
				{
					string familyName = Normalize((child == null) ? string.Empty : child.FamilyName);
					if (familyName.Length > 0)
					{
						result.Add(familyName);
					}
				}
			}
		}
		return result;
	}

	private static bool IsModelLoadableFamily(StandardLoadableFamilySnapshotItem item)
	{
		if (item == null)
		{
			return false;
		}
		return string.Equals(FamilyBrowserFamilyClassificationService.ResolveCategoryGroup(item.CategoryGroup, item.CategoryName, item.CategoryId, item.FamilyName), "Model", StringComparison.OrdinalIgnoreCase);
	}

	private static bool IsModelNestedLoadableChild(StandardNestedLoadableFamilySnapshotItem item)
	{
		if (item == null)
		{
			return false;
		}
		return string.Equals(FamilyBrowserFamilyClassificationService.ResolveCategoryGroup(item.CategoryGroup, item.CategoryName, item.CategoryId, item.FamilyName), "Model", StringComparison.OrdinalIgnoreCase);
	}

	private static void WriteWorkbook(string outputPath, string sheetName, IList<string> headers, IList<List<string>> rows)
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
		return "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><sheets><sheet name=\"" + XmlEscape(string.IsNullOrWhiteSpace(sheetName) ? "StandardList" : sheetName) + "\" sheetId=\"1\" r:id=\"rId1\"/></sheets></workbook>";
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

	private static string BuildWorksheetXml(IList<string> headers, IList<List<string>> rows)
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

	private static string BuildColumnsXml(int columnCount)
	{
		double[] widths = new double[8] { 16.0, 24.0, 54.0, 30.0, 64.0, 28.0, 42.0, 72.0 };
		int num = Math.Max(1, columnCount);
		StringBuilder builder = new StringBuilder();
		builder.Append("<cols>");
		int num2 = num;
		checked
		{
			for (int i = 1; i <= num2; i++)
			{
				double width = ((i <= widths.Length) ? widths[i - 1] : 24.0);
				builder.Append("<col min=\"" + i.ToString(CultureInfo.InvariantCulture) + "\" max=\"" + i.ToString(CultureInfo.InvariantCulture) + "\" width=\"" + width.ToString(CultureInfo.InvariantCulture) + "\" customWidth=\"1\"/>");
			}
			builder.Append("</cols>");
			return builder.ToString();
		}
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

	private static string Normalize(string value)
	{
		if (value == null)
		{
			return string.Empty;
		}
		return value.Trim().ToLowerInvariant();
	}
}
