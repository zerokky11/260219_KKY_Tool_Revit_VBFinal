using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading;
using System.Xml;
using System.Xml.Linq;
using Microsoft.VisualBasic.CompilerServices;

public sealed class FamilyBrowserStandardListService
{
	private sealed class StandardListBaselineState
	{
		public string BaselineCreatedAtUtc { get; set; }

		public string BaselineCreatedBy { get; set; }

		public string BaselineSourceSnapshotPath { get; set; }

		public int BaselineSystemExclusionVersion { get; set; }

		public List<FamilyBrowserStandardListJsonEntry> BaselineExcludedLoadableFamilies { get; set; }

		public List<FamilyBrowserStandardListJsonEntry> BaselineExcludedSystemTypes { get; set; }

		public StandardListBaselineState()
		{
			BaselineCreatedAtUtc = string.Empty;
			BaselineCreatedBy = string.Empty;
			BaselineSourceSnapshotPath = string.Empty;
			BaselineExcludedLoadableFamilies = new List<FamilyBrowserStandardListJsonEntry>();
			BaselineExcludedSystemTypes = new List<FamilyBrowserStandardListJsonEntry>();
		}
	}

	private sealed class BaselineLoadableEntryIndex
	{
		public readonly Dictionary<string, List<FamilyBrowserStandardListEntry>> EntriesByFamilyToken;

		public readonly List<FamilyBrowserStandardListEntry> WildcardFamilyEntries;

		public BaselineLoadableEntryIndex()
		{
			EntriesByFamilyToken = new Dictionary<string, List<FamilyBrowserStandardListEntry>>(StringComparer.Ordinal);
			WildcardFamilyEntries = new List<FamilyBrowserStandardListEntry>();
		}
	}

	private sealed class BaselineSystemEntryIndex
	{
		public readonly Dictionary<string, List<FamilyBrowserStandardListEntry>> EntriesByTypeToken;

		public readonly List<FamilyBrowserStandardListEntry> WildcardTypeEntries;

		public BaselineSystemEntryIndex()
		{
			EntriesByTypeToken = new Dictionary<string, List<FamilyBrowserStandardListEntry>>(StringComparer.Ordinal);
			WildcardTypeEntries = new List<FamilyBrowserStandardListEntry>();
		}
	}

	[Serializable]
	[CompilerGenerated]
	internal sealed class _Closure_0024__
	{
		public static readonly _Closure_0024__ _0024I;

		public static Func<FamilyBrowserStandardListEntry, FamilyBrowserStandardListJsonEntry> _0024I12_002D0;

		public static Func<StandardLoadableFamilySnapshotItem, bool> _0024I15_002D0;

		public static Func<StandardLoadableFamilySnapshotItem, string> _0024I15_002D1;

		public static Func<StandardLoadableFamilySnapshotItem, string> _0024I15_002D2;

		public static Func<StandardSystemTypeSnapshotItem, bool> _0024I15_002D3;

		public static Func<StandardSystemTypeSnapshotItem, string> _0024I15_002D4;

		public static Func<StandardSystemTypeSnapshotItem, string> _0024I15_002D5;

		public static Func<StandardSystemTypeSnapshotItem, string> _0024I15_002D6;

		public static Func<FamilyBrowserStandardListJsonEntry, bool> _0024I25_002D0;

		public static Func<FamilyBrowserStandardListJsonEntry, FamilyBrowserStandardListJsonEntry> _0024I25_002D1;

		public static Func<string, bool> _0024I31_002D0;

		public static Func<string, string> _0024I31_002D1;

		public static Func<FamilyBrowserStandardListJsonEntry, bool> _0024I36_002D0;

		public static Func<FamilyBrowserStandardListJsonEntry, FamilyBrowserStandardListEntry> _0024I36_002D1;

		public static Func<FamilyBrowserStandardListEntry, bool> _0024I36_002D2;

		public static Func<FamilyBrowserStandardListJsonEntry, bool> _0024I38_002D0;

		public static Func<FamilyBrowserStandardListJsonEntry, FamilyBrowserStandardListEntry> _0024I38_002D1;

		public static Func<FamilyBrowserStandardListEntry, bool> _0024I38_002D2;

		public static Func<FamilyBrowserStandardListIndexItem, bool> _0024I40_002D0;

		public static Func<FamilyBrowserStandardListIndexItem, string> _0024I40_002D1;

		public static Func<IGrouping<string, FamilyBrowserStandardListIndexItem>, FamilyBrowserStandardListIndexItem> _0024I40_002D2;

		public static Func<FamilyBrowserStandardListIndexItem, string> _0024I40_002D3;

		public static Func<FamilyBrowserStandardListIndexItem, string> _0024I40_002D4;

		public static Func<string, bool> _0024I43_002D0;

		public static Func<string, List<string>> _0024I43_002D1;

		public static Func<string, bool> _0024I44_002D0;

		public static Func<string, bool> _0024I46_002D0;

		public static Func<XElement, string> _0024I65_002D0;

		public static Func<XElement, string> _0024I68_002D0;

		public static Func<char, bool> _0024I70_002D0;

		public static Func<FamilyBrowserStandardListEntry, FamilyBrowserStandardListEntry> _0024I72_002D0;

		public static Func<FamilyBrowserStandardListEntry, bool> _0024I73_002D0;

		public static Func<FamilyBrowserStandardListEntry, FamilyBrowserStandardListEntry> _0024I73_002D1;

		static _Closure_0024__()
		{
			_0024I = new _Closure_0024__();
		}

		[SpecialName]
		internal FamilyBrowserStandardListJsonEntry _Lambda_0024__12_002D0(FamilyBrowserStandardListEntry entry)
		{
			return new FamilyBrowserStandardListJsonEntry
			{
				RowNumber = entry.RowNumber,
				Discipline = (entry.Discipline ?? string.Empty),
				Category = (entry.Category ?? string.Empty),
				Family = (entry.Family ?? string.Empty),
				TypeName = (entry.TypeName ?? string.Empty),
				Notes = (entry.Notes ?? string.Empty)
			};
		}

		[SpecialName]
		internal bool _Lambda_0024__15_002D0(StandardLoadableFamilySnapshotItem x)
		{
			return x != null;
		}

		[SpecialName]
		internal string _Lambda_0024__15_002D1(StandardLoadableFamilySnapshotItem x)
		{
			return FoldToken(x.CategoryName);
		}

		[SpecialName]
		internal string _Lambda_0024__15_002D2(StandardLoadableFamilySnapshotItem x)
		{
			return FoldToken(x.FamilyName);
		}

		[SpecialName]
		internal bool _Lambda_0024__15_002D3(StandardSystemTypeSnapshotItem x)
		{
			return x != null;
		}

		[SpecialName]
		internal string _Lambda_0024__15_002D4(StandardSystemTypeSnapshotItem x)
		{
			return FoldToken(x.CategoryName);
		}

		[SpecialName]
		internal string _Lambda_0024__15_002D5(StandardSystemTypeSnapshotItem x)
		{
			return FoldToken(x.TypeClassName);
		}

		[SpecialName]
		internal string _Lambda_0024__15_002D6(StandardSystemTypeSnapshotItem x)
		{
			return FoldToken(x.TypeName);
		}

		[SpecialName]
		internal bool _Lambda_0024__25_002D0(FamilyBrowserStandardListJsonEntry x)
		{
			return x != null;
		}

		[SpecialName]
		internal FamilyBrowserStandardListJsonEntry _Lambda_0024__25_002D1(FamilyBrowserStandardListJsonEntry x)
		{
			return new FamilyBrowserStandardListJsonEntry
			{
				RowNumber = x.RowNumber,
				Discipline = (x.Discipline ?? string.Empty),
				Category = (x.Category ?? string.Empty),
				Family = (x.Family ?? string.Empty),
				TypeName = (x.TypeName ?? string.Empty),
				Notes = (x.Notes ?? string.Empty)
			};
		}

		[SpecialName]
		internal bool _Lambda_0024__31_002D0(string filePath)
		{
			return IsSupportedStandardListExtension(Path.GetExtension(filePath));
		}

		[SpecialName]
		internal string _Lambda_0024__31_002D1(string filePath)
		{
			return Path.GetFileName(filePath);
		}

		[SpecialName]
		internal bool _Lambda_0024__36_002D0(FamilyBrowserStandardListJsonEntry entry)
		{
			return entry != null;
		}

		[SpecialName]
		internal FamilyBrowserStandardListEntry _Lambda_0024__36_002D1(FamilyBrowserStandardListJsonEntry entry)
		{
			return new FamilyBrowserStandardListEntry
			{
				RowNumber = entry.RowNumber,
				Discipline = (entry.Discipline ?? string.Empty),
				Category = (entry.Category ?? string.Empty),
				Family = (entry.Family ?? string.Empty),
				TypeName = (entry.TypeName ?? string.Empty),
				Notes = (entry.Notes ?? string.Empty)
			};
		}

		[SpecialName]
		internal bool _Lambda_0024__36_002D2(FamilyBrowserStandardListEntry entry)
		{
			if (string.IsNullOrWhiteSpace(entry.Category) && string.IsNullOrWhiteSpace(entry.Family))
			{
				return !string.IsNullOrWhiteSpace(entry.TypeName);
			}
			return true;
		}

		[SpecialName]
		internal bool _Lambda_0024__38_002D0(FamilyBrowserStandardListJsonEntry x)
		{
			return x != null;
		}

		[SpecialName]
		internal FamilyBrowserStandardListEntry _Lambda_0024__38_002D1(FamilyBrowserStandardListJsonEntry x)
		{
			return new FamilyBrowserStandardListEntry
			{
				RowNumber = x.RowNumber,
				Discipline = (x.Discipline ?? string.Empty),
				Category = (x.Category ?? string.Empty),
				Family = (x.Family ?? string.Empty),
				TypeName = (x.TypeName ?? string.Empty),
				Notes = (x.Notes ?? string.Empty)
			};
		}

		[SpecialName]
		internal bool _Lambda_0024__38_002D2(FamilyBrowserStandardListEntry entry)
		{
			if (string.IsNullOrWhiteSpace(entry.Category) && string.IsNullOrWhiteSpace(entry.Family))
			{
				return !string.IsNullOrWhiteSpace(entry.TypeName);
			}
			return true;
		}

		[SpecialName]
		internal bool _Lambda_0024__40_002D0(FamilyBrowserStandardListIndexItem item)
		{
			if (item != null)
			{
				return !string.IsNullOrWhiteSpace(item.StandardListPath);
			}
			return false;
		}

		[SpecialName]
		internal string _Lambda_0024__40_002D1(FamilyBrowserStandardListIndexItem item)
		{
			return FamilyBrowserPolicyKey.Normalize(item.SlotKey) + "|" + FamilyBrowserPolicyKey.Normalize(item.Discipline);
		}

		[SpecialName]
		internal FamilyBrowserStandardListIndexItem _Lambda_0024__40_002D2(IGrouping<string, FamilyBrowserStandardListIndexItem> group)
		{
			return group.First();
		}

		[SpecialName]
		internal string _Lambda_0024__40_002D3(FamilyBrowserStandardListIndexItem item)
		{
			return FamilyBrowserPolicyKey.Normalize(item.Discipline);
		}

		[SpecialName]
		internal string _Lambda_0024__40_002D4(FamilyBrowserStandardListIndexItem item)
		{
			return FamilyBrowserPolicyKey.Normalize(item.DisplayName);
		}

		[SpecialName]
		internal bool _Lambda_0024__43_002D0(string value)
		{
			return value != null;
		}

		[SpecialName]
		internal List<string> _Lambda_0024__43_002D1(string line)
		{
			return SplitCsvLine(line);
		}

		[SpecialName]
		internal bool _Lambda_0024__44_002D0(string value)
		{
			return !string.IsNullOrWhiteSpace(value);
		}

		[SpecialName]
		internal bool _Lambda_0024__46_002D0(string value)
		{
			return string.IsNullOrWhiteSpace(value);
		}

		[SpecialName]
		internal string _Lambda_0024__65_002D0(XElement x)
		{
			return x.Value;
		}

		[SpecialName]
		internal string _Lambda_0024__68_002D0(XElement x)
		{
			return x.Value;
		}

		[SpecialName]
		internal bool _Lambda_0024__70_002D0(char ch)
		{
			return char.IsLetter(ch);
		}

		[SpecialName]
		internal FamilyBrowserStandardListEntry _Lambda_0024__72_002D0(FamilyBrowserStandardListEntry entry)
		{
			return new FamilyBrowserStandardListEntry
			{
				RowNumber = entry.RowNumber,
				Discipline = entry.Discipline,
				Category = entry.Category,
				Family = entry.Family,
				TypeName = entry.TypeName,
				Notes = entry.Notes
			};
		}

		[SpecialName]
		internal bool _Lambda_0024__73_002D0(FamilyBrowserStandardListEntry x)
		{
			return x != null;
		}

		[SpecialName]
		internal FamilyBrowserStandardListEntry _Lambda_0024__73_002D1(FamilyBrowserStandardListEntry entry)
		{
			return new FamilyBrowserStandardListEntry
			{
				RowNumber = entry.RowNumber,
				Discipline = (entry.Discipline ?? string.Empty),
				Category = (entry.Category ?? string.Empty),
				Family = (entry.Family ?? string.Empty),
				TypeName = (entry.TypeName ?? string.Empty),
				Notes = (entry.Notes ?? string.Empty)
			};
		}
	}

	[CompilerGenerated]
	internal sealed class _Closure_0024__18_002D0
	{
		public string _0024VB_0024Local_disciplineKey;

		public string _0024VB_0024Local_disciplineLabel;

		public _Closure_0024__18_002D0(_Closure_0024__18_002D0 arg0)
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
			if (entry != null)
			{
				return DisciplineMatches(entry.Discipline, _0024VB_0024Local_disciplineKey, _0024VB_0024Local_disciplineLabel);
			}
			return false;
		}
	}

	[CompilerGenerated]
	internal sealed class _Closure_0024__19_002D0
	{
		public string _0024VB_0024Local_disciplineKey;

		public string _0024VB_0024Local_disciplineLabel;

		public _Closure_0024__19_002D0(_Closure_0024__19_002D0 arg0)
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
			if (entry != null)
			{
				return DisciplineMatches(entry.Discipline, _0024VB_0024Local_disciplineKey, _0024VB_0024Local_disciplineLabel);
			}
			return false;
		}
	}

	[CompilerGenerated]
	internal sealed class _Closure_0024__66_002D0
	{
		public string _0024VB_0024Local_requestedSheetName;

		public _Closure_0024__66_002D0(_Closure_0024__66_002D0 arg0)
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

	private const int MaxStandardListColumnScanCount = 32;

	private const int BlankRowsAfterHeaderStopThreshold = 500;

	private const int CurrentSystemBaselineExclusionVersion = 2;

	private static readonly object CacheLock = RuntimeHelpers.GetObjectValue(new object());

	private static string CachedPath = string.Empty;

	private static string CachedSheetName = string.Empty;

	private static long CachedWriteTicks = -1L;

	private static FamilyBrowserStandardListCatalog CachedCatalog;

	private FamilyBrowserStandardListService()
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
			CachedCatalog = null;
		}
		finally
		{
			if (lockTaken)
			{
				Monitor.Exit(cacheLock);
			}
		}
	}

	public static FamilyBrowserStandardListCatalog LoadForSlot(FamilyBrowserStandardLibrarySlot slot, StandardLibraryRegistrationRecord registration)
	{
		FamilyBrowserStandardListCatalog catalog = new FamilyBrowserStandardListCatalog();
		string sourcePath = ResolveConfiguredPath(slot, registration);
		if (!string.IsNullOrWhiteSpace(sourcePath))
		{
			catalog.ExplicitPath = true;
		}
		catalog.SourcePath = sourcePath ?? string.Empty;
		catalog.SheetName = ((slot == null) ? string.Empty : slot.StandardListSheetName);
		FamilyBrowserStandardListCatalog LoadForSlot;
		if (string.IsNullOrWhiteSpace(catalog.SourcePath))
		{
			LoadForSlot = catalog;
		}
		else
		{
			catalog.Exists = File.Exists(catalog.SourcePath);
			if (!catalog.Exists)
			{
				catalog.LastError = "Standard list Excel was not found.";
				LoadForSlot = catalog;
			}
			else
			{
				try
				{
					if (!IsSupportedStandardListExtension(Path.GetExtension(catalog.SourcePath)))
					{
						catalog.LastError = "Standard list must be .json, .xlsx, or .csv.";
						LoadForSlot = catalog;
					}
					else
					{
						long writeTicks = File.GetLastWriteTimeUtc(catalog.SourcePath).Ticks;
						string sheetName = (catalog.SheetName ?? string.Empty).Trim();
						object cacheLock = CacheLock;
						ObjectFlowControl.CheckForSyncLockOnValueType(cacheLock);
						bool lockTaken = false;
						try
						{
							Monitor.Enter(cacheLock, ref lockTaken);
							if (string.Equals(CachedPath, catalog.SourcePath, StringComparison.OrdinalIgnoreCase) && string.Equals(CachedSheetName, sheetName, StringComparison.OrdinalIgnoreCase) && CachedWriteTicks == writeTicks && CachedCatalog != null)
							{
								LoadForSlot = CloneCatalog(CachedCatalog);
							}
							else
							{
								List<FamilyBrowserStandardListEntry> entries = (catalog.Entries = ReadEntries(catalog.SourcePath, sheetName));
								catalog.RowCount = entries.Count;
								catalog.LastLoadedUtc = DateTime.UtcNow.ToString("O", CultureInfo.InvariantCulture);
								ApplyJsonBaseline(catalog);
								CachedPath = catalog.SourcePath;
								CachedSheetName = sheetName;
								CachedWriteTicks = writeTicks;
								CachedCatalog = CloneCatalog(catalog);
								LoadForSlot = catalog;
							}
						}
						finally
						{
							if (lockTaken)
							{
								Monitor.Exit(cacheLock);
							}
						}
					}
				}
				catch (Exception ex)
				{
					ProjectData.SetProjectError(ex);
					Exception ex2 = ex;
					catalog.LastError = ex2.Message;
					LoadForSlot = catalog;
					ProjectData.ClearProjectError();
				}
			}
		}
		return LoadForSlot;
	}

	public static FamilyBrowserStandardListMaterializeResult MaterializeJsonFromExcel(string sourcePath, string sheetName, string workspaceRoot, FamilyBrowserStandardLibrarySlot slot, string currentUser, StandardLibrarySnapshot standardSnapshot = null, string disciplineKey = "", string disciplineLabel = "")
	{
		FamilyBrowserStandardPolicyStore.RequireManagedDataRootForWrite(workspaceRoot, FamilyBrowserLanguageService.Text("Save standard family list JSON", "표준 패밀리 목록 JSON 저장"));
		string expandedSourcePath = Environment.ExpandEnvironmentVariables((sourcePath ?? string.Empty).Trim());
		if (string.IsNullOrWhiteSpace(expandedSourcePath))
		{
			throw new ArgumentException(FamilyBrowserLanguageService.Text("A standard list Excel or CSV path is required.", "표준 목록 Excel 또는 CSV 경로가 필요합니다."), "sourcePath");
		}
		if (!File.Exists(expandedSourcePath))
		{
			throw new FileNotFoundException(FamilyBrowserLanguageService.Text("Standard list Excel or CSV was not found.", "표준 목록 Excel 또는 CSV 파일을 찾지 못했습니다."), expandedSourcePath);
		}
		string extension = Path.GetExtension(expandedSourcePath);
		if (!string.Equals(extension, ".xlsx", StringComparison.OrdinalIgnoreCase) && !string.Equals(extension, ".csv", StringComparison.OrdinalIgnoreCase))
		{
			throw new InvalidDataException("Standard list source must be .xlsx or .csv.");
		}
		string effectiveSheetName = (sheetName ?? string.Empty).Trim();
		List<FamilyBrowserStandardListEntry> entries = ReadEntries(expandedSourcePath, effectiveSheetName);
		string outputPath = ResolveJsonOutputPath(workspaceRoot, slot, expandedSourcePath);
		StandardListBaselineState standardListBaselineState = TryReadExistingBaseline(outputPath);
		string effectiveDisciplineLabel = ((!string.IsNullOrWhiteSpace(disciplineLabel)) ? disciplineLabel : ((slot == null) ? string.Empty : slot.DisplayName));
		StandardListBaselineState baseline = standardListBaselineState;
		if (baseline == null)
		{
			baseline = BuildInitialBaselineExclusions(entries, standardSnapshot, disciplineKey, effectiveDisciplineLabel, currentUser);
		}
		else if (baseline.BaselineSystemExclusionVersion < 2)
		{
			StandardListBaselineState rebuiltSystemBaseline = BuildInitialBaselineExclusions(entries, standardSnapshot, disciplineKey, effectiveDisciplineLabel, currentUser);
			if (rebuiltSystemBaseline != null)
			{
				baseline.BaselineExcludedSystemTypes = rebuiltSystemBaseline.BaselineExcludedSystemTypes;
				baseline.BaselineSystemExclusionVersion = rebuiltSystemBaseline.BaselineSystemExclusionVersion;
				if (string.IsNullOrWhiteSpace(baseline.BaselineCreatedAtUtc))
				{
					baseline.BaselineCreatedAtUtc = rebuiltSystemBaseline.BaselineCreatedAtUtc;
				}
				if (string.IsNullOrWhiteSpace(baseline.BaselineCreatedBy))
				{
					baseline.BaselineCreatedBy = rebuiltSystemBaseline.BaselineCreatedBy;
				}
				if (string.IsNullOrWhiteSpace(baseline.BaselineSourceSnapshotPath))
				{
					baseline.BaselineSourceSnapshotPath = rebuiltSystemBaseline.BaselineSourceSnapshotPath;
				}
			}
		}
		FamilyBrowserStandardListJsonDocument document = new FamilyBrowserStandardListJsonDocument
		{
			SourcePath = expandedSourcePath,
			SourceSheetName = effectiveSheetName,
			GeneratedAtUtc = DateTime.UtcNow.ToString("O", CultureInfo.InvariantCulture),
			GeneratedBy = (currentUser ?? string.Empty),
			RowCount = entries.Count,
			Entries = entries.Select([SpecialName] (FamilyBrowserStandardListEntry entry) => new FamilyBrowserStandardListJsonEntry
			{
				RowNumber = entry.RowNumber,
				Discipline = (entry.Discipline ?? string.Empty),
				Category = (entry.Category ?? string.Empty),
				Family = (entry.Family ?? string.Empty),
				TypeName = (entry.TypeName ?? string.Empty),
				Notes = (entry.Notes ?? string.Empty)
			}).ToList(),
			BaselineCreatedAtUtc = ((baseline == null) ? string.Empty : baseline.BaselineCreatedAtUtc),
			BaselineCreatedBy = ((baseline == null) ? string.Empty : baseline.BaselineCreatedBy),
			BaselineSourceSnapshotPath = ((baseline == null) ? string.Empty : baseline.BaselineSourceSnapshotPath),
			BaselineSystemExclusionVersion = (baseline?.BaselineSystemExclusionVersion ?? 0),
			BaselineExcludedLoadableFamilies = ((baseline == null) ? new List<FamilyBrowserStandardListJsonEntry>() : baseline.BaselineExcludedLoadableFamilies),
			BaselineExcludedSystemTypes = ((baseline == null) ? new List<FamilyBrowserStandardListJsonEntry>() : baseline.BaselineExcludedSystemTypes)
		};
		string? directoryName = Path.GetDirectoryName(outputPath);
		if (string.IsNullOrWhiteSpace(directoryName))
		{
			throw new InvalidOperationException(FamilyBrowserLanguageService.Text("Standard list JSON output path must include a folder.", "표준 목록 JSON 출력 경로에는 폴더가 포함되어야 합니다."));
		}
		Directory.CreateDirectory(directoryName);
		string text = outputPath + ".tmp";
		File.WriteAllText(text, PlainJsonReportWriter.Serialize(document), Encoding.UTF8);
		if (File.Exists(outputPath))
		{
			File.Delete(outputPath);
		}
		File.Move(text, outputPath);
		return new FamilyBrowserStandardListMaterializeResult
		{
			SourcePath = expandedSourcePath,
			OutputPath = outputPath,
			SheetName = effectiveSheetName,
			RowCount = entries.Count
		};
	}

	private static StandardListBaselineState TryReadExistingBaseline(string outputPath)
	{
		StandardListBaselineState TryReadExistingBaseline;
		if (string.IsNullOrWhiteSpace(outputPath) || !File.Exists(outputPath))
		{
			TryReadExistingBaseline = null;
		}
		else
		{
			try
			{
				FamilyBrowserStandardListJsonDocument document = DataContractJsonFileStore.Load<FamilyBrowserStandardListJsonDocument>(outputPath);
				TryReadExistingBaseline = ((document == null) ? null : ((!string.IsNullOrWhiteSpace(document.BaselineCreatedAtUtc) || (document.BaselineExcludedLoadableFamilies != null && document.BaselineExcludedLoadableFamilies.Count > 0) || (document.BaselineExcludedSystemTypes != null && document.BaselineExcludedSystemTypes.Count > 0)) ? new StandardListBaselineState
				{
					BaselineCreatedAtUtc = (document.BaselineCreatedAtUtc ?? string.Empty),
					BaselineCreatedBy = (document.BaselineCreatedBy ?? string.Empty),
					BaselineSourceSnapshotPath = (document.BaselineSourceSnapshotPath ?? string.Empty),
					BaselineSystemExclusionVersion = document.BaselineSystemExclusionVersion,
					BaselineExcludedLoadableFamilies = CloneJsonEntries(document.BaselineExcludedLoadableFamilies),
					BaselineExcludedSystemTypes = CloneJsonEntries(document.BaselineExcludedSystemTypes)
				} : null));
			}
			catch (Exception projectError)
			{
				ProjectData.SetProjectError(projectError);
				TryReadExistingBaseline = null;
				ProjectData.ClearProjectError();
			}
		}
		return TryReadExistingBaseline;
	}

	private static StandardListBaselineState BuildInitialBaselineExclusions(IEnumerable<FamilyBrowserStandardListEntry> entries, StandardLibrarySnapshot standardSnapshot, string disciplineKey, string disciplineLabel, string currentUser)
	{
		if (standardSnapshot == null)
		{
			return null;
		}
		FamilyBrowserStandardListCatalog catalog = new FamilyBrowserStandardListCatalog
		{
			Entries = (entries ?? Enumerable.Empty<FamilyBrowserStandardListEntry>()).ToList()
		};
		if (!IsFilteringEnabled(catalog))
		{
			return null;
		}
		StandardListBaselineState baseline = new StandardListBaselineState
		{
			BaselineCreatedAtUtc = DateTime.UtcNow.ToString("O", CultureInfo.InvariantCulture),
			BaselineCreatedBy = (currentUser ?? string.Empty),
			BaselineSourceSnapshotPath = (standardSnapshot.ResolvedPath ?? string.Empty),
			BaselineSystemExclusionVersion = 2
		};
		BaselineLoadableEntryIndex loadableIndex = BuildBaselineLoadableEntryIndex(catalog, disciplineKey, disciplineLabel);
		BaselineSystemEntryIndex systemIndex = BuildBaselineSystemEntryIndex(catalog, disciplineKey, disciplineLabel);
		HashSet<string> nestedFamilyNames = BuildNestedLoadableNameSet(standardSnapshot);
		foreach (StandardLoadableFamilySnapshotItem item in (standardSnapshot.LoadableFamilies ?? new List<StandardLoadableFamilySnapshotItem>()).Where([SpecialName] (StandardLoadableFamilySnapshotItem x) => x != null).OrderBy<StandardLoadableFamilySnapshotItem, string>([SpecialName] (StandardLoadableFamilySnapshotItem x) => FoldToken(x.CategoryName), StringComparer.Ordinal).ThenBy<StandardLoadableFamilySnapshotItem, string>([SpecialName] (StandardLoadableFamilySnapshotItem x) => FoldToken(x.FamilyName), StringComparer.Ordinal))
		{
			if (!ShouldSkipLoadableFamily(item, nestedFamilyNames) && !BaselineAllowsLoadable(loadableIndex, item.CategoryName, item.FamilyName))
			{
				baseline.BaselineExcludedLoadableFamilies.Add(new FamilyBrowserStandardListJsonEntry
				{
					Discipline = (disciplineLabel ?? string.Empty),
					Category = (item.CategoryName ?? string.Empty),
					Family = (item.FamilyName ?? string.Empty),
					TypeName = string.Empty,
					Notes = "Excluded by initial standard list baseline."
				});
			}
		}
		foreach (StandardSystemTypeSnapshotItem item2 in (standardSnapshot.SystemTypes ?? new List<StandardSystemTypeSnapshotItem>()).Where([SpecialName] (StandardSystemTypeSnapshotItem x) => x != null).OrderBy<StandardSystemTypeSnapshotItem, string>([SpecialName] (StandardSystemTypeSnapshotItem x) => FoldToken(x.CategoryName), StringComparer.Ordinal).ThenBy<StandardSystemTypeSnapshotItem, string>([SpecialName] (StandardSystemTypeSnapshotItem x) => FoldToken(x.TypeClassName), StringComparer.Ordinal)
			.ThenBy<StandardSystemTypeSnapshotItem, string>([SpecialName] (StandardSystemTypeSnapshotItem x) => FoldToken(x.TypeName), StringComparer.Ordinal))
		{
			if (!BaselineAllowsSystem(systemIndex, item2.CategoryName, item2.TypeClassName, item2.TypeName))
			{
				baseline.BaselineExcludedSystemTypes.Add(new FamilyBrowserStandardListJsonEntry
				{
					Discipline = (disciplineLabel ?? string.Empty),
					Category = (item2.CategoryName ?? string.Empty),
					Family = (item2.TypeClassName ?? string.Empty),
					TypeName = (item2.TypeName ?? string.Empty),
					Notes = "Excluded by initial standard list baseline."
				});
			}
		}
		return baseline;
	}

	private static BaselineLoadableEntryIndex BuildBaselineLoadableEntryIndex(FamilyBrowserStandardListCatalog catalog, string disciplineKey, string disciplineLabel)
	{
		_Closure_0024__18_002D0 arg = default(_Closure_0024__18_002D0);
		_Closure_0024__18_002D0 CS_0024_003C_003E8__locals6 = new _Closure_0024__18_002D0(arg);
		CS_0024_003C_003E8__locals6._0024VB_0024Local_disciplineKey = disciplineKey;
		CS_0024_003C_003E8__locals6._0024VB_0024Local_disciplineLabel = disciplineLabel;
		BaselineLoadableEntryIndex index = new BaselineLoadableEntryIndex();
		if (catalog == null || catalog.Entries == null)
		{
			return index;
		}
		bool hasDisciplineSpecificRows = catalog.Entries.Any([SpecialName] (FamilyBrowserStandardListEntry familyBrowserStandardListEntry) => familyBrowserStandardListEntry != null && DisciplineMatches(familyBrowserStandardListEntry.Discipline, CS_0024_003C_003E8__locals6._0024VB_0024Local_disciplineKey, CS_0024_003C_003E8__locals6._0024VB_0024Local_disciplineLabel));
		foreach (FamilyBrowserStandardListEntry entry in catalog.Entries)
		{
			if (entry == null || !EntryAppliesToDisciplineFast(entry, CS_0024_003C_003E8__locals6._0024VB_0024Local_disciplineKey, CS_0024_003C_003E8__locals6._0024VB_0024Local_disciplineLabel, hasDisciplineSpecificRows))
			{
				continue;
			}
			if (string.IsNullOrWhiteSpace(entry.Family))
			{
				if (string.IsNullOrWhiteSpace(entry.TypeName))
				{
					index.WildcardFamilyEntries.Add(entry);
				}
			}
			else
			{
				AddBaselineEntry(index.EntriesByFamilyToken, FoldToken(entry.Family), entry);
			}
		}
		return index;
	}

	private static BaselineSystemEntryIndex BuildBaselineSystemEntryIndex(FamilyBrowserStandardListCatalog catalog, string disciplineKey, string disciplineLabel)
	{
		_Closure_0024__19_002D0 arg = default(_Closure_0024__19_002D0);
		_Closure_0024__19_002D0 CS_0024_003C_003E8__locals6 = new _Closure_0024__19_002D0(arg);
		CS_0024_003C_003E8__locals6._0024VB_0024Local_disciplineKey = disciplineKey;
		CS_0024_003C_003E8__locals6._0024VB_0024Local_disciplineLabel = disciplineLabel;
		BaselineSystemEntryIndex index = new BaselineSystemEntryIndex();
		if (catalog == null || catalog.Entries == null)
		{
			return index;
		}
		bool hasDisciplineSpecificRows = catalog.Entries.Any([SpecialName] (FamilyBrowserStandardListEntry familyBrowserStandardListEntry) => familyBrowserStandardListEntry != null && DisciplineMatches(familyBrowserStandardListEntry.Discipline, CS_0024_003C_003E8__locals6._0024VB_0024Local_disciplineKey, CS_0024_003C_003E8__locals6._0024VB_0024Local_disciplineLabel));
		foreach (FamilyBrowserStandardListEntry entry in catalog.Entries)
		{
			if (entry != null && EntryAppliesToDisciplineFast(entry, CS_0024_003C_003E8__locals6._0024VB_0024Local_disciplineKey, CS_0024_003C_003E8__locals6._0024VB_0024Local_disciplineLabel, hasDisciplineSpecificRows) && !string.IsNullOrWhiteSpace(entry.TypeName))
			{
				AddBaselineEntry(index.EntriesByTypeToken, FoldToken(entry.TypeName), entry);
			}
		}
		return index;
	}

	private static void AddBaselineEntry(Dictionary<string, List<FamilyBrowserStandardListEntry>> map, string token, FamilyBrowserStandardListEntry entry)
	{
		if (!string.IsNullOrWhiteSpace(token) && map != null && entry != null)
		{
			List<FamilyBrowserStandardListEntry> list = null;
			if (!map.TryGetValue(token, out list) || list == null)
			{
				list = (map[token] = new List<FamilyBrowserStandardListEntry>());
			}
			list.Add(entry);
		}
	}

	private static bool EntryAppliesToDisciplineFast(FamilyBrowserStandardListEntry entry, string disciplineKey, string disciplineLabel, bool hasDisciplineSpecificRows)
	{
		if (entry == null || string.IsNullOrWhiteSpace(entry.Discipline))
		{
			return true;
		}
		if (!hasDisciplineSpecificRows)
		{
			return true;
		}
		return DisciplineMatches(entry.Discipline, disciplineKey, disciplineLabel);
	}

	private static bool BaselineAllowsLoadable(BaselineLoadableEntryIndex index, string categoryName, string familyName)
	{
		if (index == null)
		{
			return false;
		}
		List<FamilyBrowserStandardListEntry> entries = null;
		if (index.EntriesByFamilyToken.TryGetValue(FoldToken(familyName), out entries) && entries != null)
		{
			foreach (FamilyBrowserStandardListEntry entry in entries)
			{
				if (entry != null && CategoryMatches(entry.Category, categoryName))
				{
					return true;
				}
			}
		}
		foreach (FamilyBrowserStandardListEntry entry2 in index.WildcardFamilyEntries)
		{
			if (entry2 != null && CategoryMatches(entry2.Category, categoryName))
			{
				return true;
			}
		}
		return false;
	}

	private static bool BaselineAllowsSystem(BaselineSystemEntryIndex index, string categoryName, string systemFamilyKind, string typeName)
	{
		if (index == null)
		{
			return false;
		}
		List<FamilyBrowserStandardListEntry> entries = null;
		if (index.EntriesByTypeToken.TryGetValue(FoldToken(typeName), out entries) && BaselineSystemEntriesAllow(entries, categoryName, systemFamilyKind))
		{
			return true;
		}
		return false;
	}

	private static bool BaselineSystemEntriesAllow(IEnumerable<FamilyBrowserStandardListEntry> entries, string categoryName, string systemFamilyKind)
	{
		if (entries == null)
		{
			return false;
		}
		foreach (FamilyBrowserStandardListEntry entry in entries)
		{
			if (entry != null && CategoryMatches(entry.Category, categoryName) && SystemFamilyMatches(entry.Family, systemFamilyKind, categoryName))
			{
				return true;
			}
		}
		return false;
	}

	private static List<FamilyBrowserStandardListJsonEntry> CloneJsonEntries(IEnumerable<FamilyBrowserStandardListJsonEntry> items)
	{
		return (from x in items ?? Enumerable.Empty<FamilyBrowserStandardListJsonEntry>()
			where x != null
			select new FamilyBrowserStandardListJsonEntry
			{
				RowNumber = x.RowNumber,
				Discipline = (x.Discipline ?? string.Empty),
				Category = (x.Category ?? string.Empty),
				Family = (x.Family ?? string.Empty),
				TypeName = (x.TypeName ?? string.Empty),
				Notes = (x.Notes ?? string.Empty)
			}).ToList();
	}

	public static FamilyBrowserStandardListIndexResult SaveStandardListIndex(string workspaceRoot, FamilyBrowserStandardPolicy policy, string currentUser)
	{
		FamilyBrowserStandardPolicyStore.RequireManagedDataRootForWrite(workspaceRoot, FamilyBrowserLanguageService.Text("Save standard family list index", "표준 패밀리 목록 인덱스 저장"));
		string standardListFolder = FamilyBrowserStandardPolicyStore.GetStandardListFolder(workspaceRoot);
		Directory.CreateDirectory(standardListFolder);
		FamilyBrowserStandardListIndexDocument document = new FamilyBrowserStandardListIndexDocument
		{
			GeneratedAtUtc = DateTime.UtcNow.ToString("O", CultureInfo.InvariantCulture),
			GeneratedBy = (currentUser ?? string.Empty),
			Items = BuildStandardListIndexItems(policy)
		};
		string outputPath = Path.Combine(standardListFolder, "standard-list-index.json");
		string text = outputPath + ".tmp";
		File.WriteAllText(text, PlainJsonReportWriter.Serialize(document), Encoding.UTF8);
		if (File.Exists(outputPath))
		{
			File.Delete(outputPath);
		}
		File.Move(text, outputPath);
		return new FamilyBrowserStandardListIndexResult
		{
			OutputPath = outputPath,
			ItemCount = document.Items.Count
		};
	}

	public static bool IsFilteringEnabled(FamilyBrowserStandardListCatalog catalog)
	{
		if (catalog != null && catalog.Entries != null)
		{
			return catalog.Entries.Count > 0;
		}
		return false;
	}

	public static bool AllowsLoadable(FamilyBrowserStandardListCatalog catalog, string disciplineKey, string disciplineLabel, string categoryName, string familyName)
	{
		if (!IsFilteringEnabled(catalog))
		{
			return true;
		}
		return catalog.Entries.Any([SpecialName] (FamilyBrowserStandardListEntry entry) => EntryAppliesToDiscipline(catalog, entry, disciplineKey, disciplineLabel) && CategoryMatches(entry.Category, categoryName) && LoadableFamilyMatches(entry, familyName));
	}

	public static bool AllowsSystem(FamilyBrowserStandardListCatalog catalog, string disciplineKey, string disciplineLabel, string categoryName, string systemFamilyKind, string typeName)
	{
		if (!IsFilteringEnabled(catalog))
		{
			return true;
		}
		return catalog.Entries.Any([SpecialName] (FamilyBrowserStandardListEntry entry) => EntryAppliesToDiscipline(catalog, entry, disciplineKey, disciplineLabel) && CategoryMatches(entry.Category, categoryName) && SystemFamilyMatches(entry.Family, systemFamilyKind, categoryName) && SystemTypeMatches(entry.TypeName, typeName));
	}

	private static string ResolveConfiguredPath(FamilyBrowserStandardLibrarySlot slot, StandardLibraryRegistrationRecord registration)
	{
		if (slot == null || string.IsNullOrWhiteSpace(slot.StandardListPath))
		{
			return string.Empty;
		}
		string raw = Environment.ExpandEnvironmentVariables(slot.StandardListPath.Trim());
		if (Path.IsPathRooted(raw))
		{
			return raw;
		}
		string standardPath = ResolveStandardPath(slot, registration);
		if (!string.IsNullOrWhiteSpace(standardPath))
		{
			string folder = Path.GetDirectoryName(standardPath);
			if (!string.IsNullOrWhiteSpace(folder))
			{
				return Path.GetFullPath(Path.Combine(folder, raw));
			}
		}
		return Path.GetFullPath(raw);
	}

	private static string TryDiscoverListPath(FamilyBrowserStandardLibrarySlot slot, StandardLibraryRegistrationRecord registration)
	{
		string standardPath = ResolveStandardPath(slot, registration);
		if (string.IsNullOrWhiteSpace(standardPath))
		{
			return string.Empty;
		}
		string folder = Path.GetDirectoryName(standardPath);
		if (string.IsNullOrWhiteSpace(folder) || !Directory.Exists(folder))
		{
			return string.Empty;
		}
		string baseName = Path.GetFileNameWithoutExtension(standardPath);
		List<string> candidates = new List<string>
		{
			Path.Combine(folder, baseName + ".json"),
			Path.Combine(folder, baseName + ".xlsx"),
			Path.Combine(folder, baseName + ".csv"),
			Path.Combine(folder, baseName + "_list.json"),
			Path.Combine(folder, baseName + "_list.xlsx"),
			Path.Combine(folder, baseName + "_list.csv"),
			Path.Combine(folder, baseName + "-list.json"),
			Path.Combine(folder, baseName + "-list.xlsx"),
			Path.Combine(folder, baseName + "-list.csv"),
			Path.Combine(folder, baseName + " list.json"),
			Path.Combine(folder, baseName + " list.xlsx"),
			Path.Combine(folder, baseName + " list.csv"),
			Path.Combine(folder, "standard-list.json"),
			Path.Combine(folder, "standard-list.xlsx"),
			Path.Combine(folder, "standard-list.csv")
		};
		foreach (string candidate in candidates)
		{
			if (File.Exists(candidate) && LooksLikeStandardList(candidate, (slot == null) ? string.Empty : slot.StandardListSheetName))
			{
				return candidate;
			}
		}
		try
		{
			foreach (string candidate2 in (from filePath in Directory.EnumerateFiles(folder)
				where IsSupportedStandardListExtension(Path.GetExtension(filePath))
				select filePath).OrderBy<string, string>([SpecialName] (string filePath) => Path.GetFileName(filePath), StringComparer.OrdinalIgnoreCase))
			{
				if (LooksLikeStandardList(candidate2, (slot == null) ? string.Empty : slot.StandardListSheetName))
				{
					return candidate2;
				}
			}
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ProjectData.ClearProjectError();
		}
		return string.Empty;
	}

	private static string ResolveStandardPath(FamilyBrowserStandardLibrarySlot slot, StandardLibraryRegistrationRecord registration)
	{
		if (slot != null && !string.IsNullOrWhiteSpace(slot.StandardRvtPath))
		{
			return Environment.ExpandEnvironmentVariables(slot.StandardRvtPath.Trim());
		}
		if (registration != null && !string.IsNullOrWhiteSpace(registration.ResolvedPath))
		{
			return Environment.ExpandEnvironmentVariables(registration.ResolvedPath.Trim());
		}
		return string.Empty;
	}

	private static bool LooksLikeStandardList(string path, string sheetName)
	{
		bool LooksLikeStandardList;
		try
		{
			LooksLikeStandardList = ReadEntries(path, sheetName).Count > 0;
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			LooksLikeStandardList = false;
			ProjectData.ClearProjectError();
		}
		return LooksLikeStandardList;
	}

	private static bool IsSupportedStandardListExtension(string extension)
	{
		if (!string.Equals(extension, ".json", StringComparison.OrdinalIgnoreCase) && !string.Equals(extension, ".xlsx", StringComparison.OrdinalIgnoreCase))
		{
			return string.Equals(extension, ".csv", StringComparison.OrdinalIgnoreCase);
		}
		return true;
	}

	private static List<FamilyBrowserStandardListEntry> ReadEntries(string path, string sheetName)
	{
		string extension = Path.GetExtension(path);
		if (string.Equals(extension, ".json", StringComparison.OrdinalIgnoreCase))
		{
			return ReadJsonEntries(path);
		}
		if (string.Equals(extension, ".csv", StringComparison.OrdinalIgnoreCase))
		{
			return ReadCsvEntries(path);
		}
		if (string.Equals(extension, ".xlsx", StringComparison.OrdinalIgnoreCase))
		{
			return ReadXlsxEntries(path, sheetName);
		}
		throw new InvalidDataException("Standard list must be .json, .xlsx, or .csv.");
	}

	private static List<FamilyBrowserStandardListEntry> ReadJsonEntries(string path)
	{
		FamilyBrowserStandardListJsonDocument document = DataContractJsonFileStore.Load<FamilyBrowserStandardListJsonDocument>(path);
		if (document == null || document.Entries == null)
		{
			return new List<FamilyBrowserStandardListEntry>();
		}
		return (from entry in document.Entries
			where entry != null
			select new FamilyBrowserStandardListEntry
			{
				RowNumber = entry.RowNumber,
				Discipline = (entry.Discipline ?? string.Empty),
				Category = (entry.Category ?? string.Empty),
				Family = (entry.Family ?? string.Empty),
				TypeName = (entry.TypeName ?? string.Empty),
				Notes = (entry.Notes ?? string.Empty)
			} into entry
			where !string.IsNullOrWhiteSpace(entry.Category) || !string.IsNullOrWhiteSpace(entry.Family) || !string.IsNullOrWhiteSpace(entry.TypeName)
			select entry).ToList();
	}

	private static void ApplyJsonBaseline(FamilyBrowserStandardListCatalog catalog)
	{
		if (catalog == null || string.IsNullOrWhiteSpace(catalog.SourcePath) || !string.Equals(Path.GetExtension(catalog.SourcePath), ".json", StringComparison.OrdinalIgnoreCase))
		{
			return;
		}
		try
		{
			FamilyBrowserStandardListJsonDocument document = DataContractJsonFileStore.Load<FamilyBrowserStandardListJsonDocument>(catalog.SourcePath);
			if (document != null)
			{
				catalog.BaselineCreatedAtUtc = document.BaselineCreatedAtUtc ?? string.Empty;
				catalog.BaselineCreatedBy = document.BaselineCreatedBy ?? string.Empty;
				catalog.BaselineSourceSnapshotPath = document.BaselineSourceSnapshotPath ?? string.Empty;
				catalog.BaselineSystemExclusionVersion = document.BaselineSystemExclusionVersion;
				catalog.BaselineExcludedLoadableFamilies = ConvertJsonEntries(document.BaselineExcludedLoadableFamilies);
				catalog.BaselineExcludedSystemTypes = ConvertJsonEntries(document.BaselineExcludedSystemTypes);
			}
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			catalog.BaselineCreatedAtUtc = string.Empty;
			catalog.BaselineCreatedBy = string.Empty;
			catalog.BaselineSourceSnapshotPath = string.Empty;
			catalog.BaselineSystemExclusionVersion = 0;
			catalog.BaselineExcludedLoadableFamilies = new List<FamilyBrowserStandardListEntry>();
			catalog.BaselineExcludedSystemTypes = new List<FamilyBrowserStandardListEntry>();
			ProjectData.ClearProjectError();
		}
	}

	private static List<FamilyBrowserStandardListEntry> ConvertJsonEntries(IEnumerable<FamilyBrowserStandardListJsonEntry> items)
	{
		return (from x in items ?? Enumerable.Empty<FamilyBrowserStandardListJsonEntry>()
			where x != null
			select new FamilyBrowserStandardListEntry
			{
				RowNumber = x.RowNumber,
				Discipline = (x.Discipline ?? string.Empty),
				Category = (x.Category ?? string.Empty),
				Family = (x.Family ?? string.Empty),
				TypeName = (x.TypeName ?? string.Empty),
				Notes = (x.Notes ?? string.Empty)
			} into entry
			where !string.IsNullOrWhiteSpace(entry.Category) || !string.IsNullOrWhiteSpace(entry.Family) || !string.IsNullOrWhiteSpace(entry.TypeName)
			select entry).ToList();
	}

	private static string ResolveJsonOutputPath(string workspaceRoot, FamilyBrowserStandardLibrarySlot slot, string sourcePath)
	{
		string standardListFolder = FamilyBrowserStandardPolicyStore.GetStandardListFolder(workspaceRoot);
		string disciplineKey = FamilyBrowserPolicyKey.Normalize((slot == null) ? string.Empty : slot.Discipline);
		if (string.IsNullOrWhiteSpace(disciplineKey))
		{
			disciplineKey = FamilyBrowserPolicyKey.Normalize((slot == null) ? string.Empty : slot.DisplayName);
		}
		if (string.IsNullOrWhiteSpace(disciplineKey))
		{
			disciplineKey = "standard";
		}
		string baseName = Path.GetFileNameWithoutExtension(sourcePath);
		if (string.IsNullOrWhiteSpace(baseName))
		{
			baseName = "standard-list";
		}
		return Path.Combine(standardListFolder, SanitizePathSegment(disciplineKey), SanitizePathSegment(baseName) + ".json");
	}

	private static List<FamilyBrowserStandardListIndexItem> BuildStandardListIndexItems(FamilyBrowserStandardPolicy policy)
	{
		List<FamilyBrowserStandardListIndexItem> items = new List<FamilyBrowserStandardListIndexItem>();
		if (policy == null)
		{
			return items;
		}
		if (policy.IntegratedLibrary != null)
		{
			AddStandardListIndexItem(items, policy.IntegratedLibrary);
		}
		if (policy.DisciplineLibraries != null)
		{
			foreach (FamilyBrowserStandardLibrarySlot slot in policy.DisciplineLibraries)
			{
				AddStandardListIndexItem(items, slot);
			}
		}
		return (from item in items
			where item != null && !string.IsNullOrWhiteSpace(item.StandardListPath)
			group item by FamilyBrowserPolicyKey.Normalize(item.SlotKey) + "|" + FamilyBrowserPolicyKey.Normalize(item.Discipline) into @group
			select @group.First()).OrderBy<FamilyBrowserStandardListIndexItem, string>([SpecialName] (FamilyBrowserStandardListIndexItem item) => FamilyBrowserPolicyKey.Normalize(item.Discipline), StringComparer.Ordinal).ThenBy<FamilyBrowserStandardListIndexItem, string>([SpecialName] (FamilyBrowserStandardListIndexItem item) => FamilyBrowserPolicyKey.Normalize(item.DisplayName), StringComparer.Ordinal).ToList();
	}

	private static void AddStandardListIndexItem(List<FamilyBrowserStandardListIndexItem> items, FamilyBrowserStandardLibrarySlot slot)
	{
		if (items != null && slot != null && !string.IsNullOrWhiteSpace(slot.StandardListPath))
		{
			items.Add(new FamilyBrowserStandardListIndexItem
			{
				SlotKey = (slot.SlotKey ?? string.Empty),
				Discipline = (slot.Discipline ?? string.Empty),
				DisplayName = (slot.DisplayName ?? string.Empty),
				StandardListPath = (slot.StandardListPath ?? string.Empty),
				StandardRvtPath = (slot.StandardRvtPath ?? string.Empty),
				SourceId = (slot.SourceId ?? string.Empty),
				Enabled = slot.Enabled
			});
		}
	}

	private static string SanitizePathSegment(string value)
	{
		string text = (value ?? string.Empty).Trim();
		if (string.IsNullOrWhiteSpace(text))
		{
			return "standard-list";
		}
		char[] invalidFileNameChars = Path.GetInvalidFileNameChars();
		foreach (char ch in invalidFileNameChars)
		{
			text = text.Replace(ch, '_');
		}
		return text;
	}

	private static List<FamilyBrowserStandardListEntry> ReadCsvEntries(string path)
	{
		List<string> lines = (from value in File.ReadAllLines(path, Encoding.UTF8)
			where value != null
			select value).ToList();
		if (lines.Count == 0)
		{
			return new List<FamilyBrowserStandardListEntry>();
		}
		return MapRows(lines.Select([SpecialName] (string line) => SplitCsvLine(line)).ToList());
	}

	private static List<FamilyBrowserStandardListEntry> ReadXlsxEntries(string path, string sheetName)
	{
		using ZipArchive archive = ZipFile.OpenRead(path);
		List<string> sharedStrings = ReadSharedStrings(archive);
		string sheetEntryName = ResolveWorksheetEntryName(archive, sheetName);
		if (string.IsNullOrWhiteSpace(sheetEntryName))
		{
			throw new InvalidDataException("No worksheet was found in the standard list Excel.");
		}
		ZipArchiveEntry? obj = archive.GetEntry(sheetEntryName) ?? throw new InvalidDataException("Worksheet XML was not found: " + sheetEntryName);
		List<List<string>> tableRows = new List<List<string>>();
		using (Stream stream = obj.Open())
		{
			bool headerFound = false;
			int blankRowsAfterHeader = 0;
			XmlReaderSettings settings = new XmlReaderSettings
			{
				IgnoreComments = true,
				IgnoreProcessingInstructions = true
			};
			using XmlReader reader = XmlReader.Create(stream, settings);
			while (reader.Read())
			{
				if (reader.NodeType != XmlNodeType.Element || !string.Equals(reader.LocalName, "row", StringComparison.OrdinalIgnoreCase))
				{
					continue;
				}
				List<string> rowValues = ReadWorksheetRow(reader, sharedStrings);
				if (!rowValues.Any([SpecialName] (string value) => !string.IsNullOrWhiteSpace(value)))
				{
					if (headerFound)
					{
						blankRowsAfterHeader = checked(blankRowsAfterHeader + 1);
						if (blankRowsAfterHeader >= 500)
						{
							break;
						}
					}
				}
				else
				{
					blankRowsAfterHeader = 0;
					tableRows.Add(rowValues);
					if (!headerFound && TryResolveHeader(rowValues) != null)
					{
						headerFound = true;
					}
				}
			}
		}
		return MapRows(tableRows);
	}

	private static List<string> ReadWorksheetRow(XmlReader rowReader, List<string> sharedStrings)
	{
		Dictionary<int, string> values = new Dictionary<int, string>();
		if (rowReader.IsEmptyElement)
		{
			return new List<string>();
		}
		checked
		{
			using (XmlReader subtree = rowReader.ReadSubtree())
			{
				while (subtree.Read())
				{
					if (subtree.NodeType == XmlNodeType.Element && string.Equals(subtree.LocalName, "c", StringComparison.OrdinalIgnoreCase))
					{
						int index = ResolveCellColumnIndex(subtree.GetAttribute("r") ?? string.Empty);
						if (index <= 0)
						{
							index = values.Count + 1;
						}
						if (index <= 32)
						{
							values[index] = ReadCellValue(subtree, subtree.GetAttribute("t") ?? string.Empty, sharedStrings);
						}
					}
				}
			}
			if (values.Count == 0)
			{
				return new List<string>();
			}
			int num = Math.Min(values.Keys.Max(), 32);
			List<string> rowValues = new List<string>();
			int num2 = num;
			for (int col = 1; col <= num2; col++)
			{
				rowValues.Add(values.ContainsKey(col) ? values[col] : string.Empty);
			}
			return rowValues;
		}
	}

	private static List<FamilyBrowserStandardListEntry> MapRows(List<List<string>> tableRows)
	{
		List<string> header = null;
		int headerIndex = -1;
		checked
		{
			int num = tableRows.Count - 1;
			for (int i = 0; i <= num; i++)
			{
				if (TryResolveHeader(tableRows[i]) != null)
				{
					header = tableRows[i];
					headerIndex = i;
					break;
				}
			}
			if (header == null)
			{
				throw new InvalidDataException("Standard list requires headers: field, Category, Family, Type, notes.");
			}
			Dictionary<int, string> map = TryResolveHeader(header);
			List<FamilyBrowserStandardListEntry> result = new List<FamilyBrowserStandardListEntry>();
			int num2 = headerIndex + 1;
			int num3 = tableRows.Count - 1;
			for (int j = num2; j <= num3; j++)
			{
				List<string> values = tableRows[j];
				if (values == null || values.All([SpecialName] (string value) => string.IsNullOrWhiteSpace(value)))
				{
					continue;
				}
				FamilyBrowserStandardListEntry entry = new FamilyBrowserStandardListEntry
				{
					RowNumber = j + 1
				};
				int num4 = values.Count - 1;
				for (int col = 0; col <= num4; col++)
				{
					if (map.ContainsKey(col))
					{
						ApplyCell(entry, map[col], values[col]);
					}
				}
				if (!string.IsNullOrWhiteSpace(entry.Category) || !string.IsNullOrWhiteSpace(entry.Family) || !string.IsNullOrWhiteSpace(entry.TypeName))
				{
					result.Add(entry);
				}
			}
			return result;
		}
	}

	private static Dictionary<int, string> TryResolveHeader(List<string> header)
	{
		if (header == null)
		{
			return null;
		}
		Dictionary<int, string> map = new Dictionary<int, string>();
		checked
		{
			int num = header.Count - 1;
			for (int i = 0; i <= num; i++)
			{
				string kind = ResolveHeaderKind(header[i]);
				if (!string.IsNullOrWhiteSpace(kind) && !map.ContainsValue(kind))
				{
					map[i] = kind;
				}
			}
			if (!map.ContainsValue("Category") || !map.ContainsValue("Family") || !map.ContainsValue("Type"))
			{
				return null;
			}
			return map;
		}
	}

	private static string ResolveHeaderKind(string value)
	{
		string text = NormalizeHeader(value);
		switch (text)
		{
		default:
			if (Operators.CompareString(text, KoField(), TextCompare: false) != 0 && Operators.CompareString(text, KoTrade(), TextCompare: false) != 0)
			{
				switch (text)
				{
				default:
					if (Operators.CompareString(text, KoCategory(), TextCompare: false) != 0)
					{
						if (Operators.CompareString(text, "family", TextCompare: false) == 0 || Operators.CompareString(text, "familyname", TextCompare: false) == 0 || Operators.CompareString(text, KoFamily(), TextCompare: false) == 0)
						{
							return "Family";
						}
						switch (text)
						{
						default:
							if (Operators.CompareString(text, KoType(), TextCompare: false) != 0)
							{
								switch (text)
								{
								default:
									if (Operators.CompareString(text, KoNotes(), TextCompare: false) != 0)
									{
										return string.Empty;
									}
									goto case "note";
								case "note":
								case "notes":
								case "memo":
								case "remark":
								case "remarks":
									return "Notes";
								}
							}
							goto case "type";
						case "type":
						case "typename":
						case "familytype":
						case "systemtype":
							return "Type";
						}
					}
					goto case "category";
				case "category":
				case "cat":
				case "revitcategory":
					return "Category";
				}
			}
			goto case "discipline";
		case "discipline":
		case "field":
		case "trade":
			return "Discipline";
		}
	}

	private static void ApplyCell(FamilyBrowserStandardListEntry entry, string headerKind, string value)
	{
		string text = (value ?? string.Empty).Trim();
		switch (headerKind)
		{
		case "Discipline":
			entry.Discipline = text;
			break;
		case "Category":
			entry.Category = text;
			break;
		case "Family":
			entry.Family = text;
			break;
		case "Type":
			entry.TypeName = text;
			break;
		case "Notes":
			entry.Notes = text;
			break;
		}
	}

	private static bool EntryAppliesToDiscipline(FamilyBrowserStandardListCatalog catalog, FamilyBrowserStandardListEntry entry, string disciplineKey, string disciplineLabel)
	{
		if (entry == null || string.IsNullOrWhiteSpace(entry.Discipline))
		{
			return true;
		}
		if (catalog == null || catalog.Entries == null || !catalog.Entries.Any([SpecialName] (FamilyBrowserStandardListEntry candidate) => DisciplineMatches(candidate.Discipline, disciplineKey, disciplineLabel)))
		{
			return true;
		}
		return DisciplineMatches(entry.Discipline, disciplineKey, disciplineLabel);
	}

	private static bool DisciplineMatches(string entryDiscipline, string disciplineKey, string disciplineLabel)
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

	private static bool LoadableFamilyMatches(FamilyBrowserStandardListEntry entry, string familyName)
	{
		if (entry == null)
		{
			return false;
		}
		if (string.IsNullOrWhiteSpace(entry.Family))
		{
			return string.IsNullOrWhiteSpace(entry.TypeName);
		}
		return SameToken(entry.Family, familyName);
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
		return item.IsShared && IsModelLoadableFamily(item) && (item.IsNestedLoadableChild || (nestedFamilyNames?.Contains(NormalizeFamilyToken(item.FamilyName)) ?? false));
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
				string nestedFamilyName = NormalizeFamilyToken(parentItem.FamilyName);
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
					string familyName = NormalizeFamilyToken(child.FamilyName);
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

	private static string NormalizeFamilyToken(string value)
	{
		return (value ?? string.Empty).Trim().ToLowerInvariant();
	}

	private static bool SystemFamilyMatches(string entryFamily, string systemFamilyKind, string categoryName)
	{
		if (string.IsNullOrWhiteSpace(entryFamily))
		{
			return true;
		}
		return SimilarToken(entryFamily, systemFamilyKind) || SimilarToken(entryFamily, categoryName);
	}

	private static bool SystemTypeMatches(string entryTypeName, string typeName)
	{
		if (string.IsNullOrWhiteSpace(entryTypeName))
		{
			return false;
		}
		return SameToken(entryTypeName, typeName);
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

	private static string FoldToken(string value)
	{
		return (value ?? string.Empty).Trim().Replace(" ", string.Empty).Replace("_", string.Empty)
			.Replace("-", string.Empty)
			.Replace(".", string.Empty)
			.Replace("/", string.Empty)
			.Replace("\\", string.Empty)
			.ToLowerInvariant();
	}

	private static string NormalizeHeader(string value)
	{
		return FoldToken(value);
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
				result.Add(string.Concat(item.Descendants(ns + "t").Select((_Closure_0024__._0024I65_002D0 == null) ? (_Closure_0024__._0024I65_002D0 = [SpecialName] (XElement x) => x.Value) : _Closure_0024__._0024I65_002D0)));
			}
		}
		return result;
	}

	private static string ResolveWorksheetEntryName(ZipArchive archive, string requestedSheetName)
	{
		_Closure_0024__66_002D0 arg = default(_Closure_0024__66_002D0);
		_Closure_0024__66_002D0 CS_0024_003C_003E8__locals3 = new _Closure_0024__66_002D0(arg);
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

	private static string ReadCellValue(XmlReader cellReader, string typeValue, List<string> sharedStrings)
	{
		if (cellReader == null || cellReader.IsEmptyElement)
		{
			return string.Empty;
		}
		string raw = string.Empty;
		List<string> inlineParts = new List<string>();
		bool inlineString = string.Equals(typeValue, "inlineStr", StringComparison.OrdinalIgnoreCase);
		using (XmlReader subtree = cellReader.ReadSubtree())
		{
			while (subtree.Read())
			{
				if (subtree.NodeType == XmlNodeType.Element)
				{
					if (string.Equals(subtree.LocalName, "v", StringComparison.OrdinalIgnoreCase))
					{
						raw = subtree.ReadElementContentAsString();
					}
					else if (inlineString && string.Equals(subtree.LocalName, "t", StringComparison.OrdinalIgnoreCase))
					{
						inlineParts.Add(subtree.ReadElementContentAsString());
					}
				}
			}
		}
		if (inlineString)
		{
			return string.Concat(inlineParts);
		}
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

	private static FamilyBrowserStandardListCatalog CloneCatalog(FamilyBrowserStandardListCatalog source)
	{
		if (source == null)
		{
			return new FamilyBrowserStandardListCatalog();
		}
		return new FamilyBrowserStandardListCatalog
		{
			SourcePath = source.SourcePath,
			SheetName = source.SheetName,
			ExplicitPath = source.ExplicitPath,
			Exists = source.Exists,
			RowCount = source.RowCount,
			LastLoadedUtc = source.LastLoadedUtc,
			LastError = source.LastError,
			BaselineCreatedAtUtc = source.BaselineCreatedAtUtc,
			BaselineCreatedBy = source.BaselineCreatedBy,
			BaselineSourceSnapshotPath = source.BaselineSourceSnapshotPath,
			BaselineSystemExclusionVersion = source.BaselineSystemExclusionVersion,
			Entries = (source.Entries ?? new List<FamilyBrowserStandardListEntry>()).Select([SpecialName] (FamilyBrowserStandardListEntry entry) => new FamilyBrowserStandardListEntry
			{
				RowNumber = entry.RowNumber,
				Discipline = entry.Discipline,
				Category = entry.Category,
				Family = entry.Family,
				TypeName = entry.TypeName,
				Notes = entry.Notes
			}).ToList(),
			BaselineExcludedLoadableFamilies = CloneCatalogEntries(source.BaselineExcludedLoadableFamilies),
			BaselineExcludedSystemTypes = CloneCatalogEntries(source.BaselineExcludedSystemTypes)
		};
	}

	private static List<FamilyBrowserStandardListEntry> CloneCatalogEntries(IEnumerable<FamilyBrowserStandardListEntry> items)
	{
		return (from entry in items ?? Enumerable.Empty<FamilyBrowserStandardListEntry>()
			where entry != null
			select new FamilyBrowserStandardListEntry
			{
				RowNumber = entry.RowNumber,
				Discipline = (entry.Discipline ?? string.Empty),
				Category = (entry.Category ?? string.Empty),
				Family = (entry.Family ?? string.Empty),
				TypeName = (entry.TypeName ?? string.Empty),
				Notes = (entry.Notes ?? string.Empty)
			}).ToList();
	}

	private static string KoField()
	{
		return "분야";
	}

	private static string KoTrade()
	{
		return "공종";
	}

	private static string KoCategory()
	{
		return "카테고리";
	}

	private static string KoFamily()
	{
		return "패밀리";
	}

	private static string KoType()
	{
		return "타입";
	}

	private static string KoNotes()
	{
		return "비고";
	}
}
