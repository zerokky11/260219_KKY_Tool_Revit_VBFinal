using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Security.Cryptography;
using System.Text;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Microsoft.VisualBasic.CompilerServices;

public sealed class StandardLibraryRegistrationService
{
	private sealed class StandardLoadableFamilyDeepMetadata
	{
		public List<StandardNestedLoadableFamilySnapshotItem> NestedLoadableFamilies { get; set; }

		public List<StandardFamilyParameterSnapshotItem> Parameters { get; set; }

		public string ContentFingerprint { get; set; }

		public string ContentSignatureDebugPath { get; set; }

		public string ContentFingerprintFailureReason { get; set; }

		public bool ThumbnailGenerated { get; set; }

		public bool ThumbnailSkipped { get; set; }

		public StandardLoadableFamilyDeepMetadata()
		{
			NestedLoadableFamilies = new List<StandardNestedLoadableFamilySnapshotItem>();
			Parameters = new List<StandardFamilyParameterSnapshotItem>();
			ContentFingerprint = string.Empty;
			ContentSignatureDebugPath = string.Empty;
			ContentFingerprintFailureReason = string.Empty;
		}
	}

	[CompilerGenerated]
	internal sealed class _Closure_0024__5_002D0
	{
		public Document _0024VB_0024Local_doc;

		public bool _0024VB_0024Local_includeDeepLoadableContent;

		public string _0024VB_0024Local_sourceId;

		public string _0024VB_0024Local_effectiveSnapshotMode;

		public string _0024VB_0024Local_sourceFileLastWriteUtc;

		public long _0024VB_0024Local_sourceFileLength;

		public Action<int, int, string> _0024VB_0024Local_progress;

		public Func<ElementId, ElementType> _0024I2;

		public _Closure_0024__5_002D0(_Closure_0024__5_002D0 arg0)
		{
			if (arg0 != null)
			{
				_0024VB_0024Local_doc = arg0._0024VB_0024Local_doc;
				_0024VB_0024Local_includeDeepLoadableContent = arg0._0024VB_0024Local_includeDeepLoadableContent;
				_0024VB_0024Local_sourceId = arg0._0024VB_0024Local_sourceId;
				_0024VB_0024Local_effectiveSnapshotMode = arg0._0024VB_0024Local_effectiveSnapshotMode;
				_0024VB_0024Local_sourceFileLastWriteUtc = arg0._0024VB_0024Local_sourceFileLastWriteUtc;
				_0024VB_0024Local_sourceFileLength = arg0._0024VB_0024Local_sourceFileLength;
				_0024VB_0024Local_progress = arg0._0024VB_0024Local_progress;
			}
		}

		[SpecialName]
		internal ElementType _Lambda_0024__2(ElementId id)
		{
			Element element = _0024VB_0024Local_doc.GetElement(id);
			return (ElementType)(object)((element is ElementType) ? element : null);
		}

		[SpecialName]
		internal void _Lambda_0024__7(int current, int total, string message)
		{
			ReportProgress(_0024VB_0024Local_progress, checked(80 + (int)Math.Round((double)current / (double)Math.Max(1, total) * 4.0)), 100, message);
		}
	}

	[CompilerGenerated]
	internal sealed class _Closure_0024__5_002D1
	{
		public StandardLoadableFamilySnapshotItem _0024VB_0024Local_item;

		public string _0024VB_0024Local_categoryName;

		public string _0024VB_0024Local_familyName;

		public _Closure_0024__5_002D0 _0024VB_0024NonLocal__0024VB_0024Closure_2;

		public _Closure_0024__5_002D1(_Closure_0024__5_002D1 arg0)
		{
			if (arg0 != null)
			{
				_0024VB_0024Local_item = arg0._0024VB_0024Local_item;
				_0024VB_0024Local_categoryName = arg0._0024VB_0024Local_categoryName;
				_0024VB_0024Local_familyName = arg0._0024VB_0024Local_familyName;
			}
		}
	}

	[CompilerGenerated]
	internal sealed class _Closure_0024__5_002D2
	{
		public string _0024VB_0024Local_previewImagePath;

		public _Closure_0024__5_002D1 _0024VB_0024NonLocal__0024VB_0024Closure_3;

		public _Closure_0024__5_002D2(_Closure_0024__5_002D2 arg0)
		{
			if (arg0 != null)
			{
				_0024VB_0024Local_previewImagePath = arg0._0024VB_0024Local_previewImagePath;
			}
		}

		[SpecialName]
		internal bool _Lambda_0024__6(StandardLoadableFamilyDeepMetadata metadata)
		{
			if (!_0024VB_0024NonLocal__0024VB_0024Closure_3._0024VB_0024NonLocal__0024VB_0024Closure_2._0024VB_0024Local_includeDeepLoadableContent)
			{
				return true;
			}
			StandardLoadableFamilySnapshotItem candidateItem = BuildThumbnailSnapshotCandidate(_0024VB_0024NonLocal__0024VB_0024Closure_3._0024VB_0024Local_item, metadata);
			return !FamilyThumbnailPreviewService.IsCachedStandardThumbnailCurrent(_0024VB_0024Local_previewImagePath, BuildThumbnailRegistration(_0024VB_0024NonLocal__0024VB_0024Closure_3._0024VB_0024NonLocal__0024VB_0024Closure_2._0024VB_0024Local_sourceId, _0024VB_0024NonLocal__0024VB_0024Closure_3._0024VB_0024NonLocal__0024VB_0024Closure_2._0024VB_0024Local_effectiveSnapshotMode, _0024VB_0024NonLocal__0024VB_0024Closure_3._0024VB_0024NonLocal__0024VB_0024Closure_2._0024VB_0024Local_sourceFileLastWriteUtc, _0024VB_0024NonLocal__0024VB_0024Closure_3._0024VB_0024NonLocal__0024VB_0024Closure_2._0024VB_0024Local_sourceFileLength), BuildThumbnailSnapshot(_0024VB_0024NonLocal__0024VB_0024Closure_3._0024VB_0024NonLocal__0024VB_0024Closure_2._0024VB_0024Local_effectiveSnapshotMode, _0024VB_0024NonLocal__0024VB_0024Closure_3._0024VB_0024NonLocal__0024VB_0024Closure_2._0024VB_0024Local_sourceFileLastWriteUtc, _0024VB_0024NonLocal__0024VB_0024Closure_3._0024VB_0024NonLocal__0024VB_0024Closure_2._0024VB_0024Local_sourceFileLength), candidateItem, _0024VB_0024NonLocal__0024VB_0024Closure_3._0024VB_0024Local_categoryName, _0024VB_0024NonLocal__0024VB_0024Closure_3._0024VB_0024Local_familyName);
		}
	}

	[CompilerGenerated]
	internal sealed class _Closure_0024__6_002D0
	{
		public ISet<string> _0024VB_0024Local_selectedNames;

		public _Closure_0024__6_002D0(_Closure_0024__6_002D0 arg0)
		{
			if (arg0 != null)
			{
				_0024VB_0024Local_selectedNames = arg0._0024VB_0024Local_selectedNames;
			}
		}

		[SpecialName]
		internal bool _Lambda_0024__1(Family x)
		{
			return _0024VB_0024Local_selectedNames.Contains(Normalize(((Element)x).Name ?? string.Empty));
		}
	}

	[CompilerGenerated]
	internal sealed class _Closure_0024__7_002D0
	{
		public ISet<string> _0024VB_0024Local_selectedNames;

		public _Closure_0024__7_002D0(_Closure_0024__7_002D0 arg0)
		{
			if (arg0 != null)
			{
				_0024VB_0024Local_selectedNames = arg0._0024VB_0024Local_selectedNames;
			}
		}

		[SpecialName]
		internal bool _Lambda_0024__1(Family x)
		{
			return _0024VB_0024Local_selectedNames.Contains(Normalize(((Element)x).Name ?? string.Empty));
		}
	}

	private static readonly HashSet<string> AllowedSystemTypeNames = new HashSet<string>(new string[21]
	{
		"WallType", "FloorType", "RoofType", "CeilingType", "StairsType", "RailingType", "DuctType", "PipeType", "FlexDuctType", "FlexPipeType",
		"DuctSystemType", "PipingSystemType", "MechanicalSystemType", "ElectricalSystemType", "CableTrayType", "ConduitType", "WireType", "DuctInsulationType", "PipeInsulationType", "DuctLiningType",
		"MullionType"
	}, StringComparer.OrdinalIgnoreCase);

	private static readonly HashSet<string> RoutingAwareSystemTypeNames = new HashSet<string>(new string[6] { "DuctType", "PipeType", "FlexDuctType", "FlexPipeType", "CableTrayType", "ConduitType" }, StringComparer.OrdinalIgnoreCase);

	private StandardLibraryRegistrationService()
	{
	}

	public static StandardLibraryRegistrationResult Register(string workspaceRoot, Application application, string selectedPath, string currentUser, string snapshotMode = "Fast", UIApplication uiApplication = null, ISet<string> nestedLoadableScanFamilyNames = null, bool forceSnapshotRebuild = false, Action<int, int, string> progress = null)
	{
		if (string.IsNullOrWhiteSpace(selectedPath))
		{
			throw new ArgumentException(T("A standard library RVT path is required.", "표준 라이브러리 RVT 경로가 필요합니다."), "selectedPath");
		}
		string resolvedPath = Path.GetFullPath(selectedPath);
		if (!File.Exists(resolvedPath))
		{
			throw new FileNotFoundException(T("Standard library RVT was not found.", "표준 라이브러리 RVT를 찾지 못했습니다."), resolvedPath);
		}
		FamilyBrowserStandardPolicyStore.RequireManagedDataRootForWrite(workspaceRoot, T("Register standard RVT", "표준 RVT 등록"));
		string sourceKind = InferSourceKind(resolvedPath);
		string sourceId = BuildSourceId(sourceKind, resolvedPath);
		string effectiveSnapshotMode = NormalizeSnapshotMode(snapshotMode);
		FileInfo fileInfo = new FileInfo(resolvedPath);
		string sourceFileLastWriteUtc = fileInfo.LastWriteTimeUtc.ToString("O", CultureInfo.InvariantCulture);
		long sourceFileLength = fileInfo.Length;
		Document standardDoc = FindOpenDocument(application, resolvedPath);
		bool shouldClose = false;
		ReportProgress(progress, 2, 100, T("Checking reusable standard snapshot...", "재사용 가능한 표준 스냅샷 확인 중..."));
		if (standardDoc == null && !forceSnapshotRebuild)
		{
			StandardLibrarySnapshotCacheHit cachedSnapshot = StandardLibraryRegistryStore.TryFindReusableSnapshot(workspaceRoot, sourceId, resolvedPath, sourceFileLastWriteUtc, sourceFileLength, effectiveSnapshotMode);
			if (cachedSnapshot != null && cachedSnapshot.Snapshot != null)
			{
				EnsureStandardLoadableFingerprints(cachedSnapshot.Snapshot, T("Reusable Standard RVT Snapshot", "재사용 표준 RVT 스냅샷"));
				StandardLibraryRegistrationRecord cachedRegistration = new StandardLibraryRegistrationRecord
				{
					SourceId = cachedSnapshot.Snapshot.SourceId,
					DisplayName = cachedSnapshot.Snapshot.DisplayName,
					SourceKind = cachedSnapshot.Snapshot.SourceKind,
					Locator = cachedSnapshot.Snapshot.Locator,
					ResolvedPath = cachedSnapshot.Snapshot.ResolvedPath,
					SnapshotMode = cachedSnapshot.Snapshot.SnapshotMode,
					SourceFileLastWriteUtc = cachedSnapshot.Snapshot.SourceFileLastWriteUtc,
					SourceFileLength = cachedSnapshot.Snapshot.SourceFileLength,
					RegisteredAtUtc = DateTime.UtcNow.ToString("O", CultureInfo.InvariantCulture),
					RegisteredBy = (currentUser ?? string.Empty),
					LastSnapshotAtUtc = cachedSnapshot.Snapshot.CapturedAtUtc,
					LastSnapshotPath = cachedSnapshot.SnapshotPath,
					RevitVersion = cachedSnapshot.Snapshot.RevitVersion,
					Summary = cachedSnapshot.Snapshot.Summary
				};
				string cachedRegistrationPath = StandardLibraryRegistryStore.SaveActiveRegistration(workspaceRoot, cachedRegistration);
				ReportProgress(progress, 100, 100, T("Reusable standard snapshot loaded.", "재사용 가능한 표준 스냅샷을 불러왔습니다."));
				return new StandardLibraryRegistrationResult
				{
					Registration = cachedRegistration,
					Snapshot = cachedSnapshot.Snapshot,
					RegistrationPath = cachedRegistrationPath,
					SnapshotPath = cachedSnapshot.SnapshotPath
				};
			}
		}
		using FamilyThumbnailConstraintDialogGuard dialogGuard = new FamilyThumbnailConstraintDialogGuard(uiApplication);
		try
		{
			if (standardDoc == null)
			{
				ReportProgress(progress, 8, 100, T("Opening standard RVT...", "표준 RVT 여는 중..."));
				standardDoc = application.OpenDocumentFile(resolvedPath);
				shouldClose = true;
			}
			ReportProgress(progress, 12, 100, T("Scanning standard RVT content...", "표준 RVT 내용 스캔 중..."));
			StandardLibrarySnapshot snapshot = BuildSnapshot(standardDoc, workspaceRoot, sourceId, sourceKind, resolvedPath, effectiveSnapshotMode, sourceFileLastWriteUtc, sourceFileLength, dialogGuard, nestedLoadableScanFamilyNames, progress);
			ReportProgress(progress, 94, 100, T("Saving standard snapshot...", "표준 스냅샷 저장 중..."));
			string snapshotPath = StandardLibraryRegistryStore.SaveSnapshot(workspaceRoot, snapshot);
			List<FamilyThumbnailAutoConfirmedDialogRecord> dialogRecords = dialogGuard.GetRecordsSince(0);
			string diagnosticReportPath = WriteRegistrationDialogDiagnosticReport(workspaceRoot, sourceId, dialogRecords);
			StandardLibraryRegistrationRecord registration = new StandardLibraryRegistrationRecord
			{
				SourceId = snapshot.SourceId,
				DisplayName = snapshot.DisplayName,
				SourceKind = snapshot.SourceKind,
				Locator = snapshot.Locator,
				ResolvedPath = snapshot.ResolvedPath,
				SnapshotMode = snapshot.SnapshotMode,
				SourceFileLastWriteUtc = snapshot.SourceFileLastWriteUtc,
				SourceFileLength = snapshot.SourceFileLength,
				RegisteredAtUtc = DateTime.UtcNow.ToString("O", CultureInfo.InvariantCulture),
				RegisteredBy = (currentUser ?? string.Empty),
				LastSnapshotAtUtc = snapshot.CapturedAtUtc,
				LastSnapshotPath = snapshotPath,
				RevitVersion = snapshot.RevitVersion,
				Summary = snapshot.Summary
			};
			ReportProgress(progress, 98, 100, T("Saving active standard registration...", "현재 표준 등록 정보 저장 중..."));
			string registrationPath = StandardLibraryRegistryStore.SaveActiveRegistration(workspaceRoot, registration);
			ReportProgress(progress, 100, 100, T("Standard RVT registration completed.", "표준 RVT 등록이 완료되었습니다."));
			return new StandardLibraryRegistrationResult
			{
				Registration = registration,
				Snapshot = snapshot,
				RegistrationPath = registrationPath,
				SnapshotPath = snapshotPath,
				DiagnosticReportPath = diagnosticReportPath,
				AutoHandledDialogs = dialogRecords
			};
		}
		finally
		{
			if (shouldClose && standardDoc != null)
			{
				standardDoc.Close(false);
			}
		}
	}

	private static StandardLibrarySnapshot BuildSnapshot(Document doc, string workspaceRoot, string sourceId, string sourceKind, string resolvedPath, string snapshotMode, string sourceFileLastWriteUtc, long sourceFileLength, FamilyThumbnailConstraintDialogGuard dialogGuard, ISet<string> nestedLoadableScanFamilyNames, Action<int, int, string> progress)
	{
		//IL_0140: Unknown result type (might be due to invalid IL or missing references)
		//IL_0904: Unknown result type (might be due to invalid IL or missing references)
		_Closure_0024__5_002D0 arg = default(_Closure_0024__5_002D0);
		_Closure_0024__5_002D0 CS_0024_003C_003E8__locals28 = new _Closure_0024__5_002D0(arg);
		CS_0024_003C_003E8__locals28._0024VB_0024Local_doc = doc;
		CS_0024_003C_003E8__locals28._0024VB_0024Local_sourceId = sourceId;
		CS_0024_003C_003E8__locals28._0024VB_0024Local_sourceFileLastWriteUtc = sourceFileLastWriteUtc;
		CS_0024_003C_003E8__locals28._0024VB_0024Local_sourceFileLength = sourceFileLength;
		CS_0024_003C_003E8__locals28._0024VB_0024Local_progress = progress;
		CS_0024_003C_003E8__locals28._0024VB_0024Local_effectiveSnapshotMode = NormalizeSnapshotMode(snapshotMode);
		CS_0024_003C_003E8__locals28._0024VB_0024Local_includeDeepLoadableContent = string.Equals(CS_0024_003C_003E8__locals28._0024VB_0024Local_effectiveSnapshotMode, "Precise", StringComparison.OrdinalIgnoreCase);
		ISet<string> normalizedNestedLoadableScanFamilyNames = NormalizeFamilyNameSet(nestedLoadableScanFamilyNames);
		StandardLibrarySnapshot snapshot = new StandardLibrarySnapshot
		{
			SnapshotSchemaVersion = 4,
			SourceId = CS_0024_003C_003E8__locals28._0024VB_0024Local_sourceId,
			DisplayName = Path.GetFileNameWithoutExtension(resolvedPath),
			SourceKind = sourceKind,
			Locator = resolvedPath,
			ResolvedPath = resolvedPath,
			SnapshotMode = CS_0024_003C_003E8__locals28._0024VB_0024Local_effectiveSnapshotMode,
			SourceFileLastWriteUtc = (CS_0024_003C_003E8__locals28._0024VB_0024Local_sourceFileLastWriteUtc ?? string.Empty),
			SourceFileLength = CS_0024_003C_003E8__locals28._0024VB_0024Local_sourceFileLength,
			CapturedAtUtc = DateTime.UtcNow.ToString("O", CultureInfo.InvariantCulture),
			RevitVersion = CS_0024_003C_003E8__locals28._0024VB_0024Local_doc.Application.VersionNumber
		};
		Dictionary<string, string> loadableContentFingerprintCache = new Dictionary<string, string>(StringComparer.Ordinal);
		string fingerprintDebugRunFolder = FingerprintDebugSignatureStore.CreateStandardRunFolder(workspaceRoot, CS_0024_003C_003E8__locals28._0024VB_0024Local_sourceId, snapshot.DisplayName, snapshot.CapturedAtUtc);
		ReportProgress(CS_0024_003C_003E8__locals28._0024VB_0024Local_progress, 14, 100, T("Collecting loadable families...", "로더블 패밀리 수집 중..."));
		List<Family> loadableFamilies = (from Family x in (IEnumerable)new FilteredElementCollector(CS_0024_003C_003E8__locals28._0024VB_0024Local_doc).OfClass(typeof(Family))
			where FamilyBrowserFamilyClassificationService.IsBrowserLoadableFamily(x)
			select x).OrderBy<Family, string>([SpecialName] (Family x) => Normalize(((Element)x).Name), StringComparer.Ordinal).ToList();
		int loadableTotal = Math.Max(1, loadableFamilies.Count);
		checked
		{
			int num = loadableFamilies.Count - 1;
			_Closure_0024__5_002D1 closure_0024__5_002D = default(_Closure_0024__5_002D1);
			_Closure_0024__5_002D2 closure_0024__5_002D2 = default(_Closure_0024__5_002D2);
			for (int loadableIndex = 0; loadableIndex <= num; loadableIndex++)
			{
				closure_0024__5_002D = new _Closure_0024__5_002D1(closure_0024__5_002D);
				closure_0024__5_002D._0024VB_0024NonLocal__0024VB_0024Closure_2 = CS_0024_003C_003E8__locals28;
				Family family = loadableFamilies[loadableIndex];
				ReportProgress(closure_0024__5_002D._0024VB_0024NonLocal__0024VB_0024Closure_2._0024VB_0024Local_progress, 24 + (int)Math.Round((double)loadableIndex / (double)loadableTotal * 52.0), 100, T("Scanning loadable family ", "로더블 패밀리 스캔 중 ") + (loadableIndex + 1).ToString(CultureInfo.InvariantCulture) + "/" + loadableFamilies.Count.ToString(CultureInfo.InvariantCulture) + ": " + (((Element)family).Name ?? string.Empty));
				List<string> typeNames = (from x in family.GetFamilySymbolIds().Select([SpecialName] (ElementId id) =>
					{
						Element element = closure_0024__5_002D._0024VB_0024NonLocal__0024VB_0024Closure_2._0024VB_0024Local_doc.GetElement(id);
						return (ElementType)(object)((element is ElementType) ? element : null);
					})
					where x != null
					select ((Element)x).Name ?? string.Empty).OrderBy<string, string>([SpecialName] (string x) => Normalize(x), StringComparer.Ordinal).ToList();
				closure_0024__5_002D._0024VB_0024Local_categoryName = FamilyBrowserFamilyClassificationService.ResolveCategoryName(family);
				closure_0024__5_002D._0024VB_0024Local_familyName = ((Element)family).Name ?? string.Empty;
				dialogGuard?.SetCurrentFamily(closure_0024__5_002D._0024VB_0024Local_categoryName, closure_0024__5_002D._0024VB_0024Local_familyName);
				closure_0024__5_002D._0024VB_0024Local_item = null;
				int dialogRecordStart = dialogGuard?.RecordCount ?? 0;
				try
				{
					string fastContentFingerprint = string.Empty;
					string fastContentSignatureDebugPath = string.Empty;
					string fastContentFingerprintFailureReason = string.Empty;
					if (!closure_0024__5_002D._0024VB_0024NonLocal__0024VB_0024Closure_2._0024VB_0024Local_includeDeepLoadableContent)
					{
						LoadableFamilyContentSignatureResult signatureResult = LoadableFamilyContentSignatureService.BuildResult(closure_0024__5_002D._0024VB_0024NonLocal__0024VB_0024Closure_2._0024VB_0024Local_doc, family, closure_0024__5_002D._0024VB_0024NonLocal__0024VB_0024Closure_2._0024VB_0024Local_includeDeepLoadableContent);
						fastContentFingerprint = signatureResult.Fingerprint ?? string.Empty;
						if (string.IsNullOrWhiteSpace(fastContentFingerprint))
						{
							fastContentFingerprintFailureReason = ResolveFingerprintFailureReason(signatureResult, "Fast standard family fingerprint was empty.");
						}
						fastContentSignatureDebugPath = FingerprintDebugSignatureStore.SaveLoadableSignature(fingerprintDebugRunFolder, "standard", snapshot.DisplayName, closure_0024__5_002D._0024VB_0024Local_categoryName, closure_0024__5_002D._0024VB_0024Local_familyName, fastContentFingerprint, signatureResult);
						CacheLoadableContentFingerprint(closure_0024__5_002D._0024VB_0024NonLocal__0024VB_0024Closure_2._0024VB_0024Local_doc, family, loadableContentFingerprintCache, closure_0024__5_002D._0024VB_0024NonLocal__0024VB_0024Closure_2._0024VB_0024Local_includeDeepLoadableContent, fastContentFingerprint);
					}
					closure_0024__5_002D._0024VB_0024Local_item = new StandardLoadableFamilySnapshotItem
					{
						FamilyName = closure_0024__5_002D._0024VB_0024Local_familyName,
						CategoryName = closure_0024__5_002D._0024VB_0024Local_categoryName,
						CategoryId = FamilyBrowserFamilyClassificationService.ResolveCategoryId(family),
						CategoryGroup = FamilyBrowserFamilyClassificationService.ResolveCategoryGroup(family),
						MetadataMode = (closure_0024__5_002D._0024VB_0024NonLocal__0024VB_0024Closure_2._0024VB_0024Local_includeDeepLoadableContent ? "Precise" : "Fast"),
						TypeCount = typeNames.Count,
						TypeNames = typeNames,
						Parameters = CaptureLoadableFamilyParameters(closure_0024__5_002D._0024VB_0024NonLocal__0024VB_0024Closure_2._0024VB_0024Local_doc, family),
						ContentFingerprint = (closure_0024__5_002D._0024VB_0024NonLocal__0024VB_0024Closure_2._0024VB_0024Local_includeDeepLoadableContent ? string.Empty : fastContentFingerprint),
						ContentSignatureDebugPath = (closure_0024__5_002D._0024VB_0024NonLocal__0024VB_0024Closure_2._0024VB_0024Local_includeDeepLoadableContent ? string.Empty : fastContentSignatureDebugPath),
						ContentFingerprintFailureReason = (closure_0024__5_002D._0024VB_0024NonLocal__0024VB_0024Closure_2._0024VB_0024Local_includeDeepLoadableContent ? string.Empty : fastContentFingerprintFailureReason),
						UniqueId = (((Element)family).UniqueId ?? string.Empty),
						IsShared = ResolveIsShared(family)
					};
					if (ShouldCaptureLoadableFamilyDeepMetadata(closure_0024__5_002D._0024VB_0024Local_familyName, closure_0024__5_002D._0024VB_0024NonLocal__0024VB_0024Closure_2._0024VB_0024Local_includeDeepLoadableContent, normalizedNestedLoadableScanFamilyNames))
					{
						closure_0024__5_002D2 = new _Closure_0024__5_002D2(closure_0024__5_002D2);
						closure_0024__5_002D2._0024VB_0024NonLocal__0024VB_0024Closure_3 = closure_0024__5_002D;
						closure_0024__5_002D2._0024VB_0024Local_previewImagePath = (closure_0024__5_002D2._0024VB_0024NonLocal__0024VB_0024Closure_3._0024VB_0024NonLocal__0024VB_0024Closure_2._0024VB_0024Local_includeDeepLoadableContent ? BuildPreciseScanThumbnailPath(workspaceRoot, closure_0024__5_002D2._0024VB_0024NonLocal__0024VB_0024Closure_3._0024VB_0024NonLocal__0024VB_0024Closure_2._0024VB_0024Local_sourceId, closure_0024__5_002D2._0024VB_0024NonLocal__0024VB_0024Closure_3._0024VB_0024Local_categoryName, closure_0024__5_002D2._0024VB_0024NonLocal__0024VB_0024Closure_3._0024VB_0024Local_familyName) : string.Empty);
						StandardLoadableFamilyDeepMetadata deepMetadata = CaptureLoadableFamilyDeepMetadata(closure_0024__5_002D2._0024VB_0024NonLocal__0024VB_0024Closure_3._0024VB_0024NonLocal__0024VB_0024Closure_2._0024VB_0024Local_doc, family, closure_0024__5_002D2._0024VB_0024Local_previewImagePath, closure_0024__5_002D2._0024VB_0024NonLocal__0024VB_0024Closure_3._0024VB_0024NonLocal__0024VB_0024Closure_2._0024VB_0024Local_includeDeepLoadableContent, fingerprintDebugRunFolder, snapshot.DisplayName, closure_0024__5_002D2._Lambda_0024__6);
						closure_0024__5_002D2._0024VB_0024NonLocal__0024VB_0024Closure_3._0024VB_0024Local_item.NestedLoadableFamilies = deepMetadata.NestedLoadableFamilies;
						if (deepMetadata.Parameters.Count > 0)
						{
							closure_0024__5_002D2._0024VB_0024NonLocal__0024VB_0024Closure_3._0024VB_0024Local_item.Parameters = DeduplicateParameterSnapshots(closure_0024__5_002D2._0024VB_0024NonLocal__0024VB_0024Closure_3._0024VB_0024Local_item.Parameters.Concat(deepMetadata.Parameters));
						}
						if (closure_0024__5_002D2._0024VB_0024NonLocal__0024VB_0024Closure_3._0024VB_0024NonLocal__0024VB_0024Closure_2._0024VB_0024Local_includeDeepLoadableContent)
						{
							closure_0024__5_002D2._0024VB_0024NonLocal__0024VB_0024Closure_3._0024VB_0024Local_item.ContentFingerprint = deepMetadata.ContentFingerprint ?? string.Empty;
							closure_0024__5_002D2._0024VB_0024NonLocal__0024VB_0024Closure_3._0024VB_0024Local_item.ContentSignatureDebugPath = deepMetadata.ContentSignatureDebugPath ?? string.Empty;
							closure_0024__5_002D2._0024VB_0024NonLocal__0024VB_0024Closure_3._0024VB_0024Local_item.ContentFingerprintFailureReason = deepMetadata.ContentFingerprintFailureReason ?? string.Empty;
							CacheLoadableContentFingerprint(closure_0024__5_002D2._0024VB_0024NonLocal__0024VB_0024Closure_3._0024VB_0024NonLocal__0024VB_0024Closure_2._0024VB_0024Local_doc, family, loadableContentFingerprintCache, includeDeepLoadableContent: true, closure_0024__5_002D2._0024VB_0024NonLocal__0024VB_0024Closure_3._0024VB_0024Local_item.ContentFingerprint);
							if (deepMetadata.ThumbnailGenerated && !string.IsNullOrWhiteSpace(closure_0024__5_002D2._0024VB_0024Local_previewImagePath))
							{
								FamilyThumbnailPreviewService.WriteStandardThumbnailCacheMetadata(closure_0024__5_002D2._0024VB_0024Local_previewImagePath, BuildThumbnailRegistration(closure_0024__5_002D2._0024VB_0024NonLocal__0024VB_0024Closure_3._0024VB_0024NonLocal__0024VB_0024Closure_2._0024VB_0024Local_sourceId, closure_0024__5_002D2._0024VB_0024NonLocal__0024VB_0024Closure_3._0024VB_0024NonLocal__0024VB_0024Closure_2._0024VB_0024Local_effectiveSnapshotMode, closure_0024__5_002D2._0024VB_0024NonLocal__0024VB_0024Closure_3._0024VB_0024NonLocal__0024VB_0024Closure_2._0024VB_0024Local_sourceFileLastWriteUtc, closure_0024__5_002D2._0024VB_0024NonLocal__0024VB_0024Closure_3._0024VB_0024NonLocal__0024VB_0024Closure_2._0024VB_0024Local_sourceFileLength), BuildThumbnailSnapshot(closure_0024__5_002D2._0024VB_0024NonLocal__0024VB_0024Closure_3._0024VB_0024NonLocal__0024VB_0024Closure_2._0024VB_0024Local_effectiveSnapshotMode, closure_0024__5_002D2._0024VB_0024NonLocal__0024VB_0024Closure_3._0024VB_0024NonLocal__0024VB_0024Closure_2._0024VB_0024Local_sourceFileLastWriteUtc, closure_0024__5_002D2._0024VB_0024NonLocal__0024VB_0024Closure_3._0024VB_0024NonLocal__0024VB_0024Closure_2._0024VB_0024Local_sourceFileLength), closure_0024__5_002D2._0024VB_0024NonLocal__0024VB_0024Closure_3._0024VB_0024Local_item, closure_0024__5_002D2._0024VB_0024NonLocal__0024VB_0024Closure_3._0024VB_0024Local_categoryName, closure_0024__5_002D2._0024VB_0024NonLocal__0024VB_0024Closure_3._0024VB_0024Local_familyName);
							}
						}
					}
					else
					{
						closure_0024__5_002D._0024VB_0024Local_item.NestedLoadableFamilies = new List<StandardNestedLoadableFamilySnapshotItem>();
					}
				}
				finally
				{
					dialogGuard?.ClearCurrentFamily();
				}
				if (dialogGuard != null && closure_0024__5_002D._0024VB_0024Local_item != null)
				{
					ApplyDialogCancelFingerprintFailure(closure_0024__5_002D._0024VB_0024Local_item, family, dialogGuard.GetRecordsSince(dialogRecordStart), fingerprintDebugRunFolder, snapshot.DisplayName);
				}
				snapshot.LoadableFamilies.Add(closure_0024__5_002D._0024VB_0024Local_item);
			}
			ReportProgress(CS_0024_003C_003E8__locals28._0024VB_0024Local_progress, 77, 100, T("Resolving nested helper families...", "하위/보조 패밀리 판별 중..."));
			MarkNestedLoadableChildren(snapshot);
			ReportProgress(CS_0024_003C_003E8__locals28._0024VB_0024Local_progress, 80, 100, T("Building system type dependency map...", "시스템 타입 의존성 맵 작성 중..."));
			Dictionary<string, SystemTypeSemanticSnapshot> semanticMap = SystemTypeSemanticFingerprintCatalogService.BuildSnapshotMap(CS_0024_003C_003E8__locals28._0024VB_0024Local_doc, CS_0024_003C_003E8__locals28._0024VB_0024Local_sourceId, loadableContentFingerprintCache, CS_0024_003C_003E8__locals28._0024VB_0024Local_includeDeepLoadableContent, [SpecialName] (int current, int total, string message) =>
			{
				ReportProgress(CS_0024_003C_003E8__locals28._0024VB_0024Local_progress, 80 + (int)Math.Round((double)current / (double)Math.Max(1, total) * 4.0), 100, message);
			});
			ReportProgress(CS_0024_003C_003E8__locals28._0024VB_0024Local_progress, 84, 100, T("Collecting supported system types...", "지원되는 시스템 타입 수집 중..."));
			List<ElementType> systemTypes = (from ElementType x in (IEnumerable)new FilteredElementCollector(CS_0024_003C_003E8__locals28._0024VB_0024Local_doc).WhereElementIsElementType()
				where x != null
				where !(x is FamilySymbol)
				where AllowedSystemTypeNames.Contains(((object)x).GetType().Name)
				select x).OrderBy<ElementType, string>([SpecialName] (ElementType x) => ((object)x).GetType().Name, StringComparer.Ordinal).ThenBy<ElementType, string>([SpecialName] (ElementType x) => Normalize(((Element)x).Name), StringComparer.Ordinal).ToList();
			int systemTotal = Math.Max(1, systemTypes.Count);
			int num2 = systemTypes.Count - 1;
			for (int systemIndex = 0; systemIndex <= num2; systemIndex++)
			{
				ElementType systemType = systemTypes[systemIndex];
				SystemTypeSemanticSnapshot semanticSnapshot = ResolveSemanticSnapshot(semanticMap, ((object)systemType).GetType().Name, ResolveCategoryName((Element)(object)systemType), ((Element)systemType).Name ?? string.Empty);
				ReportProgress(CS_0024_003C_003E8__locals28._0024VB_0024Local_progress, 85 + (int)Math.Round((double)systemIndex / (double)systemTotal * 7.0), 100, T("Scanning system type ", "시스템 타입 스캔 중 ") + (systemIndex + 1).ToString(CultureInfo.InvariantCulture) + "/" + systemTypes.Count.ToString(CultureInfo.InvariantCulture) + ": " + ((object)systemType).GetType().Name + " / " + (((Element)systemType).Name ?? string.Empty));
				snapshot.SystemTypes.Add(new StandardSystemTypeSnapshotItem
				{
					TypeName = (((Element)systemType).Name ?? string.Empty),
					CategoryName = ResolveCategoryName((Element)(object)systemType),
					CategoryId = ResolveCategoryId((Element)(object)systemType),
					TypeClassName = ((object)systemType).GetType().Name,
					UniqueId = (((Element)systemType).UniqueId ?? string.Empty),
					SupportsRoutingDependencies = RoutingAwareSystemTypeNames.Contains(((object)systemType).GetType().Name),
					SemanticFingerprint = ((semanticSnapshot == null) ? string.Empty : SystemTypeFingerprintService.Compute(semanticSnapshot)),
					ClassificationCode = ((semanticSnapshot == null) ? string.Empty : semanticSnapshot.ClassificationCode),
					SegmentName = ((semanticSnapshot == null) ? string.Empty : semanticSnapshot.SegmentName),
					MaterialName = ((semanticSnapshot == null) ? string.Empty : semanticSnapshot.MaterialName),
					Shape = ((semanticSnapshot == null) ? string.Empty : semanticSnapshot.Shape),
					RoutingPreferenceSignature = ((semanticSnapshot == null) ? string.Empty : semanticSnapshot.RoutingPreferenceSignature),
					CompoundStructureSignature = ((semanticSnapshot == null) ? string.Empty : semanticSnapshot.CompoundStructureSignature),
					Layers = CaptureSystemTypeLayers(CS_0024_003C_003E8__locals28._0024VB_0024Local_doc, systemType)
				});
			}
			ReportProgress(CS_0024_003C_003E8__locals28._0024VB_0024Local_progress, 93, 100, T("Summarizing standard snapshot...", "표준 스냅샷 요약 중..."));
			snapshot.Summary = new StandardLibrarySnapshotSummary
			{
				LoadableFamilyCount = snapshot.LoadableFamilies.Count,
				LoadableTypeCount = snapshot.LoadableFamilies.Sum([SpecialName] (StandardLoadableFamilySnapshotItem x) => x.TypeCount),
				SystemTypeCount = snapshot.SystemTypes.Count,
				SystemTypeClassCount = snapshot.SystemTypes.Select([SpecialName] (StandardSystemTypeSnapshotItem x) => x.TypeClassName).Distinct<string>(StringComparer.OrdinalIgnoreCase).Count()
			};
			EnsureStandardLoadableFingerprints(snapshot, T("Standard RVT Scan", "표준 RVT 스캔"));
			return snapshot;
		}
	}

	public static StandardLibraryPartialRefreshResult RefreshSelectedLoadableFamilies(string workspaceRoot, Application application, StandardLibraryRegistrationRecord registration, string currentUser, ISet<string> familyNames, UIApplication uiApplication = null, Action<int, int, string> progress = null)
	{
		//IL_01a1: Unknown result type (might be due to invalid IL or missing references)
		_Closure_0024__6_002D0 arg = default(_Closure_0024__6_002D0);
		_Closure_0024__6_002D0 CS_0024_003C_003E8__locals5 = new _Closure_0024__6_002D0(arg);
		if (registration == null)
		{
			throw new ArgumentNullException("registration");
		}
		FamilyBrowserStandardPolicyStore.RequireManagedDataRootForWrite(workspaceRoot, T("Refresh selected standard families", "선택 표준 패밀리 갱신"));
		CS_0024_003C_003E8__locals5._0024VB_0024Local_selectedNames = NormalizeFamilyNameSet(familyNames);
		if (CS_0024_003C_003E8__locals5._0024VB_0024Local_selectedNames.Count == 0)
		{
			throw new ArgumentException(T("At least one family name is required.", "패밀리 이름을 하나 이상 선택해야 합니다."), "familyNames");
		}
		string resolvedPath = ((!string.IsNullOrWhiteSpace(registration.ResolvedPath)) ? registration.ResolvedPath : registration.Locator);
		if (string.IsNullOrWhiteSpace(resolvedPath))
		{
			throw new InvalidOperationException(T("The registered standard RVT path is empty.", "등록된 표준 RVT 경로가 비어 있습니다."));
		}
		resolvedPath = Path.GetFullPath(resolvedPath);
		if (!File.Exists(resolvedPath))
		{
			throw new FileNotFoundException(T("Standard library RVT was not found.", "표준 라이브러리 RVT를 찾지 못했습니다."), resolvedPath);
		}
		if (string.IsNullOrWhiteSpace(registration.LastSnapshotPath) || !File.Exists(registration.LastSnapshotPath))
		{
			throw new InvalidOperationException(T("The existing standard snapshot is missing. Run a fast standard scan first.", "기존 표준 스냅샷이 없습니다. 먼저 표준 빠른 스캔을 실행하세요."));
		}
		StandardLibrarySnapshot snapshot = DataContractJsonFileStore.Load<StandardLibrarySnapshot>(registration.LastSnapshotPath);
		if (snapshot == null)
		{
			throw new InvalidOperationException(T("The existing standard snapshot could not be loaded. Run a fast standard scan first.", "기존 표준 스냅샷을 불러올 수 없습니다. 먼저 표준 빠른 스캔을 실행하세요."));
		}
		string sourceKind = InferSourceKind(resolvedPath);
		string sourceId = BuildSourceId(sourceKind, resolvedPath);
		FileInfo fileInfo = new FileInfo(resolvedPath);
		string sourceFileLastWriteUtc = fileInfo.LastWriteTimeUtc.ToString("O", CultureInfo.InvariantCulture);
		long sourceFileLength = fileInfo.Length;
		Document standardDoc = FindOpenDocument(application, resolvedPath);
		bool shouldClose = false;
		StandardLibraryPartialRefreshResult result = new StandardLibraryPartialRefreshResult
		{
			RequestedCount = CS_0024_003C_003E8__locals5._0024VB_0024Local_selectedNames.Count
		};
		ReportProgress(progress, 2, 100, T("Opening standard RVT for selected family refresh...", "선택 패밀리 갱신을 위해 표준 RVT 여는 중..."));
		checked
		{
			using FamilyThumbnailConstraintDialogGuard dialogGuard = new FamilyThumbnailConstraintDialogGuard(uiApplication);
			try
			{
				if (standardDoc == null)
				{
					standardDoc = application.OpenDocumentFile(resolvedPath);
					shouldClose = true;
				}
				List<Family> families = (from Family x in (IEnumerable)new FilteredElementCollector(standardDoc).OfClass(typeof(Family))
					where FamilyBrowserFamilyClassificationService.IsBrowserLoadableFamily(x)
					where CS_0024_003C_003E8__locals5._0024VB_0024Local_selectedNames.Contains(Normalize(((Element)x).Name ?? string.Empty))
					select x).OrderBy<Family, string>([SpecialName] (Family x) => Normalize(FamilyBrowserFamilyClassificationService.ResolveCategoryName(x)) + "|" + Normalize(((Element)x).Name), StringComparer.Ordinal).ToList();
				HashSet<string> foundNames = new HashSet<string>(StringComparer.Ordinal);
				foreach (Family family in families)
				{
					foundNames.Add(Normalize(((Element)family).Name ?? string.Empty));
				}
				foreach (string selectedName in CS_0024_003C_003E8__locals5._0024VB_0024Local_selectedNames)
				{
					if (!foundNames.Contains(selectedName))
					{
						result.MissingFamilyNames.Add(selectedName);
					}
				}
				Dictionary<string, string> loadableContentFingerprintCache = new Dictionary<string, string>(StringComparer.Ordinal);
				string fingerprintDebugRunFolder = FingerprintDebugSignatureStore.CreateStandardRunFolder(workspaceRoot, sourceId, ResolveStandardDebugContextName(standardDoc, sourceId), DateTime.UtcNow.ToString("O", CultureInfo.InvariantCulture));
				int total = Math.Max(1, families.Count);
				int num = families.Count - 1;
				for (int familyIndex = 0; familyIndex <= num; familyIndex++)
				{
					Family family2 = families[familyIndex];
					string categoryName = FamilyBrowserFamilyClassificationService.ResolveCategoryName(family2);
					string familyName = ((Element)family2).Name ?? string.Empty;
					ReportProgress(progress, 8 + (int)Math.Round((double)familyIndex / (double)total * 82.0), 100, T("Refreshing selected family ", "선택 패밀리 갱신 중 ") + (familyIndex + 1).ToString(CultureInfo.InvariantCulture) + "/" + families.Count.ToString(CultureInfo.InvariantCulture) + ": " + familyName);
					int dialogRecordStart = dialogGuard.RecordCount;
					dialogGuard.SetCurrentFamily(categoryName, familyName);
					StandardLoadableFamilySnapshotItem item = null;
					try
					{
						item = BuildLoadableFamilySnapshotItem(standardDoc, family2, loadableContentFingerprintCache, includeDeepLoadableContent: true, forceDeepMetadata: true, workspaceRoot, sourceId, fingerprintDebugRunFolder, ResolveStandardDebugContextName(standardDoc, sourceId));
						ApplyDialogCancelFingerprintFailure(item, family2, dialogGuard.GetRecordsSince(dialogRecordStart), fingerprintDebugRunFolder, ResolveStandardDebugContextName(standardDoc, sourceId));
						EnsureStandardLoadableFingerprint(item, T("Selected Precise Refresh", "선택 패밀리 정밀 갱신"));
						MergeLoadableFamilySnapshotItem(snapshot, item);
						result.UpdatedCount++;
					}
					finally
					{
						dialogGuard.ClearCurrentFamily();
					}
					List<FamilyThumbnailAutoConfirmedDialogRecord> newDialogRecords = dialogGuard.GetRecordsSince(dialogRecordStart);
					if (newDialogRecords.Count > 0)
					{
						result.AutoHandledDialogs.AddRange(newDialogRecords);
					}
				}
				ReportProgress(progress, 92, 100, T("Updating nested family markers...", "하위 패밀리 표시 갱신 중..."));
				MarkNestedLoadableChildren(snapshot);
				RefreshSnapshotHeaderAndSummary(snapshot, sourceId, sourceKind, resolvedPath, sourceFileLastWriteUtc, sourceFileLength, standardDoc.Application.VersionNumber);
				ReportProgress(progress, 96, 100, T("Saving merged standard snapshot...", "병합된 표준 스냅샷 저장 중..."));
				string snapshotPath = StandardLibraryRegistryStore.SaveSnapshot(workspaceRoot, snapshot);
				string diagnosticReportPath = WriteRegistrationDialogDiagnosticReport(workspaceRoot, sourceId, result.AutoHandledDialogs);
				StandardLibraryRegistrationRecord updatedRegistration = BuildRegistrationFromSnapshot(snapshot, snapshotPath, currentUser);
				string registrationPath = StandardLibraryRegistryStore.SaveActiveRegistration(workspaceRoot, updatedRegistration);
				ReportProgress(progress, 100, 100, T("Selected standard families refreshed.", "선택한 표준 패밀리 갱신이 완료되었습니다."));
				result.Registration = updatedRegistration;
				result.Snapshot = snapshot;
				result.SnapshotPath = snapshotPath;
				result.RegistrationPath = registrationPath;
				result.DiagnosticReportPath = diagnosticReportPath;
				return result;
			}
			finally
			{
				if (shouldClose && standardDoc != null)
				{
					standardDoc.Close(false);
				}
			}
		}
	}

	public static StandardLibraryPartialRefreshResult RefreshSelectedLoadableNestedMetadata(string workspaceRoot, Application application, StandardLibraryRegistrationRecord registration, string currentUser, ISet<string> familyNames, UIApplication uiApplication = null, Action<int, int, string> progress = null)
	{
		//IL_01a1: Unknown result type (might be due to invalid IL or missing references)
		_Closure_0024__7_002D0 arg = default(_Closure_0024__7_002D0);
		_Closure_0024__7_002D0 CS_0024_003C_003E8__locals5 = new _Closure_0024__7_002D0(arg);
		if (registration == null)
		{
			throw new ArgumentNullException("registration");
		}
		FamilyBrowserStandardPolicyStore.RequireManagedDataRootForWrite(workspaceRoot, T("Refresh nested family metadata", "하위 패밀리 메타데이터 갱신"));
		CS_0024_003C_003E8__locals5._0024VB_0024Local_selectedNames = NormalizeFamilyNameSet(familyNames);
		if (CS_0024_003C_003E8__locals5._0024VB_0024Local_selectedNames.Count == 0)
		{
			throw new ArgumentException(T("At least one family name is required.", "패밀리 이름을 하나 이상 선택해야 합니다."), "familyNames");
		}
		string resolvedPath = ((!string.IsNullOrWhiteSpace(registration.ResolvedPath)) ? registration.ResolvedPath : registration.Locator);
		if (string.IsNullOrWhiteSpace(resolvedPath))
		{
			throw new InvalidOperationException(T("The registered standard RVT path is empty.", "등록된 표준 RVT 경로가 비어 있습니다."));
		}
		resolvedPath = Path.GetFullPath(resolvedPath);
		if (!File.Exists(resolvedPath))
		{
			throw new FileNotFoundException(T("Standard library RVT was not found.", "표준 라이브러리 RVT를 찾지 못했습니다."), resolvedPath);
		}
		if (string.IsNullOrWhiteSpace(registration.LastSnapshotPath) || !File.Exists(registration.LastSnapshotPath))
		{
			throw new InvalidOperationException(T("The existing standard snapshot is missing. Run a fast standard scan first.", "기존 표준 스냅샷이 없습니다. 먼저 표준 빠른 스캔을 실행하세요."));
		}
		StandardLibrarySnapshot snapshot = DataContractJsonFileStore.Load<StandardLibrarySnapshot>(registration.LastSnapshotPath);
		if (snapshot == null)
		{
			throw new InvalidOperationException(T("The existing standard snapshot could not be loaded. Run a fast standard scan first.", "기존 표준 스냅샷을 불러올 수 없습니다. 먼저 표준 빠른 스캔을 실행하세요."));
		}
		string sourceKind = InferSourceKind(resolvedPath);
		string sourceId = BuildSourceId(sourceKind, resolvedPath);
		FileInfo fileInfo = new FileInfo(resolvedPath);
		string sourceFileLastWriteUtc = fileInfo.LastWriteTimeUtc.ToString("O", CultureInfo.InvariantCulture);
		long sourceFileLength = fileInfo.Length;
		Document standardDoc = FindOpenDocument(application, resolvedPath);
		bool shouldClose = false;
		StandardLibraryPartialRefreshResult result = new StandardLibraryPartialRefreshResult
		{
			RequestedCount = CS_0024_003C_003E8__locals5._0024VB_0024Local_selectedNames.Count
		};
		ReportProgress(progress, 2, 100, T("Opening standard RVT for nested family metadata refresh...", "하위 패밀리 메타데이터 갱신을 위해 표준 RVT 여는 중..."));
		checked
		{
			using FamilyThumbnailConstraintDialogGuard dialogGuard = new FamilyThumbnailConstraintDialogGuard(uiApplication);
			try
			{
				if (standardDoc == null)
				{
					standardDoc = application.OpenDocumentFile(resolvedPath);
					shouldClose = true;
				}
				List<Family> families = (from Family x in (IEnumerable)new FilteredElementCollector(standardDoc).OfClass(typeof(Family))
					where FamilyBrowserFamilyClassificationService.IsBrowserLoadableFamily(x)
					where CS_0024_003C_003E8__locals5._0024VB_0024Local_selectedNames.Contains(Normalize(((Element)x).Name ?? string.Empty))
					select x).OrderBy<Family, string>([SpecialName] (Family x) => Normalize(FamilyBrowserFamilyClassificationService.ResolveCategoryName(x)) + "|" + Normalize(((Element)x).Name), StringComparer.Ordinal).ToList();
				HashSet<string> foundNames = new HashSet<string>(StringComparer.Ordinal);
				foreach (Family family in families)
				{
					foundNames.Add(Normalize(((Element)family).Name ?? string.Empty));
				}
				foreach (string selectedName in CS_0024_003C_003E8__locals5._0024VB_0024Local_selectedNames)
				{
					if (!foundNames.Contains(selectedName))
					{
						result.MissingFamilyNames.Add(selectedName);
					}
				}
				int total = Math.Max(1, families.Count);
				int num = families.Count - 1;
				for (int familyIndex = 0; familyIndex <= num; familyIndex++)
				{
					Family family2 = families[familyIndex];
					string categoryName = FamilyBrowserFamilyClassificationService.ResolveCategoryName(family2);
					string familyName = ((Element)family2).Name ?? string.Empty;
					ReportProgress(progress, 8 + (int)Math.Round((double)familyIndex / (double)total * 82.0), 100, T("Refreshing nested family metadata ", "하위 패밀리 메타데이터 갱신 중 ") + (familyIndex + 1).ToString(CultureInfo.InvariantCulture) + "/" + families.Count.ToString(CultureInfo.InvariantCulture) + ": " + familyName);
					int dialogRecordStart = dialogGuard.RecordCount;
					dialogGuard.SetCurrentFamily(categoryName, familyName);
					try
					{
						StandardLoadableFamilyDeepMetadata metadata = CaptureLoadableFamilyDeepMetadata(standardDoc, family2);
						MergeLoadableFamilyNestedMetadata(snapshot, family2, metadata);
						result.UpdatedCount++;
					}
					finally
					{
						dialogGuard.ClearCurrentFamily();
					}
					List<FamilyThumbnailAutoConfirmedDialogRecord> newDialogRecords = dialogGuard.GetRecordsSince(dialogRecordStart);
					if (newDialogRecords.Count > 0)
					{
						result.AutoHandledDialogs.AddRange(newDialogRecords);
					}
				}
				ReportProgress(progress, 92, 100, T("Updating nested family markers...", "하위 패밀리 표시 갱신 중..."));
				MarkNestedLoadableChildren(snapshot);
				RefreshSnapshotHeaderAndSummary(snapshot, sourceId, sourceKind, resolvedPath, sourceFileLastWriteUtc, sourceFileLength, standardDoc.Application.VersionNumber);
				ReportProgress(progress, 96, 100, T("Saving nested family metadata...", "하위 패밀리 메타데이터 저장 중..."));
				string snapshotPath = StandardLibraryRegistryStore.SaveSnapshot(workspaceRoot, snapshot);
				string diagnosticReportPath = WriteRegistrationDialogDiagnosticReport(workspaceRoot, sourceId, result.AutoHandledDialogs);
				StandardLibraryRegistrationRecord updatedRegistration = BuildRegistrationFromSnapshot(snapshot, snapshotPath, currentUser);
				string registrationPath = StandardLibraryRegistryStore.SaveActiveRegistration(workspaceRoot, updatedRegistration);
				ReportProgress(progress, 100, 100, T("Nested family metadata refreshed.", "하위 패밀리 메타데이터 갱신이 완료되었습니다."));
				result.Registration = updatedRegistration;
				result.Snapshot = snapshot;
				result.SnapshotPath = snapshotPath;
				result.RegistrationPath = registrationPath;
				result.DiagnosticReportPath = diagnosticReportPath;
				return result;
			}
			finally
			{
				if (shouldClose && standardDoc != null)
				{
					standardDoc.Close(false);
				}
			}
		}
	}

	private static StandardLoadableFamilySnapshotItem BuildLoadableFamilySnapshotItem(Document doc, Family family, IDictionary<string, string> loadableContentFingerprintCache, bool includeDeepLoadableContent, bool forceDeepMetadata, string workspaceRoot = "", string sourceId = "", string fingerprintDebugRunFolder = "", string fingerprintDebugContextName = "")
	{
		List<string> typeNames = (from x in family.GetFamilySymbolIds().Select([SpecialName] (ElementId id) =>
			{
				Element element = doc.GetElement(id);
				return (ElementType)(object)((element is ElementType) ? element : null);
			})
			where x != null
			select ((Element)x).Name ?? string.Empty).OrderBy<string, string>([SpecialName] (string x) => Normalize(x), StringComparer.Ordinal).ToList();
		string familyName = ((Element)family).Name ?? string.Empty;
		StandardLoadableFamilySnapshotItem item = new StandardLoadableFamilySnapshotItem
		{
			FamilyName = familyName,
			CategoryName = FamilyBrowserFamilyClassificationService.ResolveCategoryName(family),
			CategoryId = FamilyBrowserFamilyClassificationService.ResolveCategoryId(family),
			CategoryGroup = FamilyBrowserFamilyClassificationService.ResolveCategoryGroup(family),
			MetadataMode = ((includeDeepLoadableContent || forceDeepMetadata) ? "Precise" : "Fast"),
			TypeCount = typeNames.Count,
			TypeNames = typeNames,
			Parameters = CaptureLoadableFamilyParameters(doc, family),
			UniqueId = (((Element)family).UniqueId ?? string.Empty),
			IsShared = ResolveIsShared(family)
		};
		string debugContextName = (string.IsNullOrWhiteSpace(fingerprintDebugContextName) ? ResolveStandardDebugContextName(doc, sourceId) : fingerprintDebugContextName);
		if (!includeDeepLoadableContent && !forceDeepMetadata)
		{
			LoadableFamilyContentSignatureResult signatureResult = LoadableFamilyContentSignatureService.BuildResult(doc, family, includeDeepLoadableContent);
			item.ContentFingerprint = signatureResult.Fingerprint ?? string.Empty;
			if (string.IsNullOrWhiteSpace(item.ContentFingerprint))
			{
				item.ContentFingerprintFailureReason = ResolveFingerprintFailureReason(signatureResult, "Standard family fingerprint was empty.");
			}
			item.ContentSignatureDebugPath = FingerprintDebugSignatureStore.SaveLoadableSignature(fingerprintDebugRunFolder, "standard", debugContextName, item.CategoryName, familyName, item.ContentFingerprint, signatureResult);
			CacheLoadableContentFingerprint(doc, family, loadableContentFingerprintCache, includeDeepLoadableContent, item.ContentFingerprint);
		}
		else
		{
			item.ContentFingerprint = string.Empty;
			item.ContentSignatureDebugPath = string.Empty;
		}
		if (includeDeepLoadableContent || forceDeepMetadata)
		{
			string previewImagePath = ((includeDeepLoadableContent || forceDeepMetadata) ? BuildPreciseScanThumbnailPath(workspaceRoot, sourceId, item.CategoryName, familyName) : string.Empty);
			StandardLoadableFamilyDeepMetadata deepMetadata = CaptureLoadableFamilyDeepMetadata(doc, family, previewImagePath, includeDeepLoadableContent || forceDeepMetadata, fingerprintDebugRunFolder, debugContextName);
			item.NestedLoadableFamilies = deepMetadata.NestedLoadableFamilies;
			if (deepMetadata.Parameters.Count > 0)
			{
				item.Parameters = DeduplicateParameterSnapshots(item.Parameters.Concat(deepMetadata.Parameters));
			}
			item.ContentFingerprint = deepMetadata.ContentFingerprint ?? string.Empty;
			item.ContentSignatureDebugPath = deepMetadata.ContentSignatureDebugPath ?? string.Empty;
			item.ContentFingerprintFailureReason = deepMetadata.ContentFingerprintFailureReason ?? string.Empty;
			CacheLoadableContentFingerprint(doc, family, loadableContentFingerprintCache, includeDeepLoadableContent: true, item.ContentFingerprint);
		}
		else
		{
			item.NestedLoadableFamilies = new List<StandardNestedLoadableFamilySnapshotItem>();
		}
		return item;
	}

	private static void MergeLoadableFamilySnapshotItem(StandardLibrarySnapshot snapshot, StandardLoadableFamilySnapshotItem item)
	{
		if (snapshot != null && item != null)
		{
			if (snapshot.LoadableFamilies == null)
			{
				snapshot.LoadableFamilies = new List<StandardLoadableFamilySnapshotItem>();
			}
			string text = BuildNestedLoadableKey(item);
			snapshot.LoadableFamilies.RemoveAll([SpecialName] (StandardLoadableFamilySnapshotItem existing) => existing != null && (string.Equals(BuildNestedLoadableKey(existing), text, StringComparison.Ordinal) || (string.IsNullOrWhiteSpace(text) && string.Equals(Normalize(existing.FamilyName), Normalize(item.FamilyName), StringComparison.Ordinal))));
			snapshot.LoadableFamilies.Add(item);
			snapshot.LoadableFamilies = snapshot.LoadableFamilies.OrderBy<StandardLoadableFamilySnapshotItem, string>([SpecialName] (StandardLoadableFamilySnapshotItem x) => Normalize(x.CategoryName) + "|" + Normalize(x.FamilyName), StringComparer.Ordinal).ToList();
		}
	}

	private static void MergeLoadableFamilyNestedMetadata(StandardLibrarySnapshot snapshot, Family family, StandardLoadableFamilyDeepMetadata metadata)
	{
		if (snapshot == null || family == null)
		{
			return;
		}
		if (snapshot.LoadableFamilies == null)
		{
			snapshot.LoadableFamilies = new List<StandardLoadableFamilySnapshotItem>();
		}
		string categoryName = FamilyBrowserFamilyClassificationService.ResolveCategoryName(family);
		string text = ((Element)family).Name ?? string.Empty;
		string b = BuildNestedLoadableKey(FamilyBrowserFamilyClassificationService.ResolveCategoryId(family), categoryName, text);
		StandardLoadableFamilySnapshotItem existing = snapshot.LoadableFamilies.FirstOrDefault([SpecialName] (StandardLoadableFamilySnapshotItem x) => string.Equals(BuildNestedLoadableKey(x), b, StringComparison.Ordinal) || string.Equals(Normalize(x.FamilyName), Normalize(text), StringComparison.Ordinal));
		if (existing == null)
		{
			List<string> typeNames = (from x in family.GetFamilySymbolIds().Select([SpecialName] (ElementId id) =>
				{
					Element element = ((Element)family).Document.GetElement(id);
					return (ElementType)(object)((element is ElementType) ? element : null);
				})
				where x != null
				select ((Element)x).Name ?? string.Empty).OrderBy<string, string>([SpecialName] (string x) => Normalize(x), StringComparer.Ordinal).ToList();
			existing = new StandardLoadableFamilySnapshotItem
			{
				FamilyName = text,
				CategoryName = categoryName,
				CategoryId = FamilyBrowserFamilyClassificationService.ResolveCategoryId(family),
				CategoryGroup = FamilyBrowserFamilyClassificationService.ResolveCategoryGroup(family),
				MetadataMode = "Fast",
				TypeCount = typeNames.Count,
				TypeNames = typeNames,
				Parameters = CaptureLoadableFamilyParameters(((Element)family).Document, family),
				UniqueId = (((Element)family).UniqueId ?? string.Empty),
				IsShared = ResolveIsShared(family)
			};
			snapshot.LoadableFamilies.Add(existing);
		}
		existing.NestedLoadableFamilies = (metadata?.NestedLoadableFamilies ?? new List<StandardNestedLoadableFamilySnapshotItem>()).ToList();
		if (!string.Equals(existing.MetadataMode, "Precise", StringComparison.OrdinalIgnoreCase))
		{
			existing.MetadataMode = "Fast+Nested";
		}
		if (metadata != null && metadata.Parameters != null && metadata.Parameters.Count > 0)
		{
			existing.Parameters = DeduplicateParameterSnapshots((existing.Parameters ?? new List<StandardFamilyParameterSnapshotItem>()).Concat(metadata.Parameters));
		}
	}

	private static void CacheLoadableContentFingerprint(Document doc, Family family, IDictionary<string, string> loadableContentFingerprintCache, bool includeDeepLoadableContent, string contentFingerprint)
	{
		if (doc != null && family != null && loadableContentFingerprintCache != null)
		{
			string cacheKey = LoadableFamilyContentSignatureService.BuildCacheKey(doc, family, includeDeepLoadableContent);
			if (!string.IsNullOrWhiteSpace(cacheKey))
			{
				loadableContentFingerprintCache[cacheKey] = contentFingerprint ?? string.Empty;
			}
		}
	}

	private static string ResolveStandardDebugContextName(Document doc, string sourceId)
	{
		try
		{
			if (doc != null && !string.IsNullOrWhiteSpace(doc.PathName))
			{
				return Path.GetFileNameWithoutExtension(doc.PathName);
			}
			if (doc != null && !string.IsNullOrWhiteSpace(doc.Title))
			{
				return doc.Title;
			}
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ProjectData.ClearProjectError();
		}
		return sourceId ?? string.Empty;
	}

	private static StandardLoadableFamilySnapshotItem BuildThumbnailSnapshotCandidate(StandardLoadableFamilySnapshotItem item, StandardLoadableFamilyDeepMetadata metadata)
	{
		if (item == null)
		{
			return null;
		}
		List<StandardFamilyParameterSnapshotItem> metadataParameters = metadata?.Parameters ?? new List<StandardFamilyParameterSnapshotItem>();
		StandardLoadableFamilySnapshotItem standardLoadableFamilySnapshotItem = new StandardLoadableFamilySnapshotItem();
		standardLoadableFamilySnapshotItem.FamilyName = item.FamilyName;
		standardLoadableFamilySnapshotItem.CategoryName = item.CategoryName;
		standardLoadableFamilySnapshotItem.CategoryId = item.CategoryId;
		standardLoadableFamilySnapshotItem.CategoryGroup = item.CategoryGroup;
		standardLoadableFamilySnapshotItem.MetadataMode = "Precise";
		standardLoadableFamilySnapshotItem.TypeCount = item.TypeCount;
		standardLoadableFamilySnapshotItem.TypeNames = (item.TypeNames ?? new List<string>()).ToList();
		standardLoadableFamilySnapshotItem.Parameters = DeduplicateParameterSnapshots((item.Parameters ?? new List<StandardFamilyParameterSnapshotItem>()).Concat(metadataParameters));
		standardLoadableFamilySnapshotItem.ContentFingerprint = metadata?.ContentFingerprint ?? item.ContentFingerprint;
		standardLoadableFamilySnapshotItem.ContentSignatureDebugPath = metadata?.ContentSignatureDebugPath ?? item.ContentSignatureDebugPath;
		standardLoadableFamilySnapshotItem.UniqueId = item.UniqueId;
		standardLoadableFamilySnapshotItem.IsShared = item.IsShared;
		standardLoadableFamilySnapshotItem.IsNestedLoadableChild = item.IsNestedLoadableChild;
		standardLoadableFamilySnapshotItem.NestedLoadableFamilies = (metadata?.NestedLoadableFamilies ?? new List<StandardNestedLoadableFamilySnapshotItem>()).ToList();
		return standardLoadableFamilySnapshotItem;
	}

	private static StandardLibraryRegistrationRecord BuildThumbnailRegistration(string sourceId, string snapshotMode, string sourceFileLastWriteUtc, long sourceFileLength)
	{
		return new StandardLibraryRegistrationRecord
		{
			SourceId = (sourceId ?? string.Empty),
			SnapshotMode = (snapshotMode ?? string.Empty),
			SourceFileLastWriteUtc = (sourceFileLastWriteUtc ?? string.Empty),
			SourceFileLength = sourceFileLength
		};
	}

	private static StandardLibrarySnapshot BuildThumbnailSnapshot(string snapshotMode, string sourceFileLastWriteUtc, long sourceFileLength)
	{
		return new StandardLibrarySnapshot
		{
			SnapshotMode = (snapshotMode ?? string.Empty),
			SourceFileLastWriteUtc = (sourceFileLastWriteUtc ?? string.Empty),
			SourceFileLength = sourceFileLength
		};
	}

	private static void RefreshSnapshotHeaderAndSummary(StandardLibrarySnapshot snapshot, string sourceId, string sourceKind, string resolvedPath, string sourceFileLastWriteUtc, long sourceFileLength, string revitVersion)
	{
		if (snapshot != null)
		{
			snapshot.SnapshotSchemaVersion = Math.Max(snapshot.SnapshotSchemaVersion, 4);
			snapshot.SourceId = sourceId;
			snapshot.DisplayName = Path.GetFileNameWithoutExtension(resolvedPath);
			snapshot.SourceKind = sourceKind;
			snapshot.Locator = resolvedPath;
			snapshot.ResolvedPath = resolvedPath;
			snapshot.SourceFileLastWriteUtc = sourceFileLastWriteUtc ?? string.Empty;
			snapshot.SourceFileLength = sourceFileLength;
			snapshot.CapturedAtUtc = DateTime.UtcNow.ToString("O", CultureInfo.InvariantCulture);
			snapshot.RevitVersion = revitVersion ?? string.Empty;
			if (string.Equals(snapshot.SnapshotMode, "Precise", StringComparison.OrdinalIgnoreCase))
			{
				snapshot.SnapshotMode = "Precise";
			}
			else
			{
				snapshot.SnapshotMode = "Fast";
			}
			snapshot.Summary = new StandardLibrarySnapshotSummary
			{
				LoadableFamilyCount = (snapshot.LoadableFamilies ?? new List<StandardLoadableFamilySnapshotItem>()).Count,
				LoadableTypeCount = (snapshot.LoadableFamilies ?? new List<StandardLoadableFamilySnapshotItem>()).Sum([SpecialName] (StandardLoadableFamilySnapshotItem x) => x?.TypeCount ?? 0),
				SystemTypeCount = (snapshot.SystemTypes ?? new List<StandardSystemTypeSnapshotItem>()).Count,
				SystemTypeClassCount = (from x in snapshot.SystemTypes ?? new List<StandardSystemTypeSnapshotItem>()
					where x != null
					select x.TypeClassName).Distinct<string>(StringComparer.OrdinalIgnoreCase).Count()
			};
		}
	}

	private static StandardLibraryRegistrationRecord BuildRegistrationFromSnapshot(StandardLibrarySnapshot snapshot, string snapshotPath, string currentUser)
	{
		return new StandardLibraryRegistrationRecord
		{
			SourceId = snapshot.SourceId,
			DisplayName = snapshot.DisplayName,
			SourceKind = snapshot.SourceKind,
			Locator = snapshot.Locator,
			ResolvedPath = snapshot.ResolvedPath,
			SnapshotMode = snapshot.SnapshotMode,
			SourceFileLastWriteUtc = snapshot.SourceFileLastWriteUtc,
			SourceFileLength = snapshot.SourceFileLength,
			RegisteredAtUtc = DateTime.UtcNow.ToString("O", CultureInfo.InvariantCulture),
			RegisteredBy = (currentUser ?? string.Empty),
			LastSnapshotAtUtc = snapshot.CapturedAtUtc,
			LastSnapshotPath = snapshotPath,
			RevitVersion = snapshot.RevitVersion,
			Summary = snapshot.Summary
		};
	}

	private static void ReportProgress(Action<int, int, string> progress, int current, int total, string message)
	{
		if (progress != null)
		{
			try
			{
				int safeTotal = Math.Max(1, total);
				progress(Math.Max(0, Math.Min(current, safeTotal)), safeTotal, message ?? string.Empty);
			}
			catch (Exception projectError)
			{
				ProjectData.SetProjectError(projectError);
				ProjectData.ClearProjectError();
			}
		}
	}

	private static string T(string englishText, string koreanText)
	{
		return FamilyBrowserLanguageService.Text(englishText, koreanText);
	}

	private static bool ShouldCaptureLoadableFamilyDeepMetadata(string familyName, bool includeDeepLoadableContent, ISet<string> nestedLoadableScanFamilyNames)
	{
		if (includeDeepLoadableContent)
		{
			return true;
		}
		if (nestedLoadableScanFamilyNames == null || string.IsNullOrWhiteSpace(familyName))
		{
			return false;
		}
		return nestedLoadableScanFamilyNames.Contains(Normalize(familyName));
	}

	private static StandardLoadableFamilyDeepMetadata CaptureLoadableFamilyDeepMetadata(Document doc, Family parentFamily, string previewImagePath = "", bool captureContentFingerprint = false, string fingerprintDebugRunFolder = "", string fingerprintDebugContextName = "", Func<StandardLoadableFamilyDeepMetadata, bool> shouldGenerateThumbnail = null)
	{
		//IL_00eb: Unknown result type (might be due to invalid IL or missing references)
		StandardLoadableFamilyDeepMetadata metadata = new StandardLoadableFamilyDeepMetadata();
		Dictionary<string, StandardNestedLoadableFamilySnapshotItem> result = new Dictionary<string, StandardNestedLoadableFamilySnapshotItem>(StringComparer.Ordinal);
		StandardLoadableFamilyDeepMetadata CaptureLoadableFamilyDeepMetadata;
		if (doc == null || parentFamily == null)
		{
			CaptureLoadableFamilyDeepMetadata = metadata;
		}
		else
		{
			string parentKey = BuildNestedLoadableKey(FamilyBrowserFamilyClassificationService.ResolveCategoryId(parentFamily), FamilyBrowserFamilyClassificationService.ResolveCategoryName(parentFamily), ((Element)parentFamily).Name ?? string.Empty);
			Document familyDoc = null;
			try
			{
				familyDoc = doc.EditFamily(parentFamily);
				if (familyDoc == null)
				{
					if (captureContentFingerprint)
					{
						RecordDeepMetadataSignatureFailure(metadata, parentFamily, fingerprintDebugRunFolder, fingerprintDebugContextName, "EditFamily returned no document.");
					}
					CaptureLoadableFamilyDeepMetadata = metadata;
					goto IL_02cf;
				}
				if (captureContentFingerprint)
				{
					LoadableFamilyContentSignatureResult signatureResult = LoadableFamilyContentSignatureService.BuildResultFromOpenFamilyDocument(parentFamily, familyDoc);
					metadata.ContentFingerprint = signatureResult.Fingerprint ?? string.Empty;
					if (string.IsNullOrWhiteSpace(metadata.ContentFingerprint))
					{
						metadata.ContentFingerprintFailureReason = ResolveFingerprintFailureReason(signatureResult, "Open standard family document fingerprint was empty.");
					}
					metadata.ContentSignatureDebugPath = FingerprintDebugSignatureStore.SaveLoadableSignature(fingerprintDebugRunFolder, "standard", fingerprintDebugContextName, FamilyBrowserFamilyClassificationService.ResolveCategoryName(parentFamily), ((Element)parentFamily).Name ?? string.Empty, metadata.ContentFingerprint, signatureResult);
				}
				metadata.Parameters = CaptureFamilyDocumentParameters(familyDoc);
				foreach (FamilyInstance instance in ((IEnumerable)new FilteredElementCollector(familyDoc).OfClass(typeof(FamilyInstance))).Cast<FamilyInstance>())
				{
					TryAddNestedLoadableFamily(result, familyDoc, parentKey, ResolveNestedFamily(instance));
				}
				bool generateThumbnail = !string.IsNullOrWhiteSpace(previewImagePath);
				if (generateThumbnail && shouldGenerateThumbnail != null)
				{
					try
					{
						generateThumbnail = shouldGenerateThumbnail(metadata);
					}
					catch (Exception projectError)
					{
						ProjectData.SetProjectError(projectError);
						generateThumbnail = true;
						ProjectData.ClearProjectError();
					}
				}
				if (generateThumbnail)
				{
					metadata.ThumbnailGenerated = TryGeneratePreciseScanThumbnail(familyDoc, previewImagePath);
				}
				else if (!string.IsNullOrWhiteSpace(previewImagePath))
				{
					metadata.ThumbnailSkipped = true;
				}
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				if (captureContentFingerprint)
				{
					RecordDeepMetadataSignatureFailure(metadata, parentFamily, fingerprintDebugRunFolder, fingerprintDebugContextName, "EditFamily/deep metadata failed: " + ex2.GetType().Name + " - " + ex2.Message);
				}
				metadata.NestedLoadableFamilies = result.Values.OrderBy<StandardNestedLoadableFamilySnapshotItem, string>([SpecialName] (StandardNestedLoadableFamilySnapshotItem x) => Normalize(x.CategoryName), StringComparer.Ordinal).ThenBy<StandardNestedLoadableFamilySnapshotItem, string>([SpecialName] (StandardNestedLoadableFamilySnapshotItem x) => Normalize(x.FamilyName), StringComparer.Ordinal).ToList();
				CaptureLoadableFamilyDeepMetadata = metadata;
				ProjectData.ClearProjectError();
				goto IL_02cf;
			}
			finally
			{
				if (familyDoc != null)
				{
					try
					{
						familyDoc.Close(false);
					}
					catch (Exception projectError2)
					{
						ProjectData.SetProjectError(projectError2);
						ProjectData.ClearProjectError();
					}
				}
			}
			metadata.NestedLoadableFamilies = result.Values.OrderBy<StandardNestedLoadableFamilySnapshotItem, string>([SpecialName] (StandardNestedLoadableFamilySnapshotItem x) => Normalize(x.CategoryName), StringComparer.Ordinal).ThenBy<StandardNestedLoadableFamilySnapshotItem, string>([SpecialName] (StandardNestedLoadableFamilySnapshotItem x) => Normalize(x.FamilyName), StringComparer.Ordinal).ToList();
			CaptureLoadableFamilyDeepMetadata = metadata;
		}
		goto IL_02cf;
		IL_02cf:
		return CaptureLoadableFamilyDeepMetadata;
	}

	private static string ResolveFingerprintFailureReason(LoadableFamilyContentSignatureResult signatureResult, string fallbackReason)
	{
		if (signatureResult != null && !string.IsNullOrWhiteSpace(signatureResult.ErrorMessage))
		{
			return signatureResult.ErrorMessage.Trim();
		}
		if (signatureResult != null && !string.IsNullOrWhiteSpace(signatureResult.Signature))
		{
			return (string.IsNullOrWhiteSpace(fallbackReason) ? "Fingerprint was empty without an error message." : fallbackReason) + " Signature diagnostics were written, but content-fingerprint was blank.";
		}
		return (string.IsNullOrWhiteSpace(fallbackReason) ? "Fingerprint was empty without an error message." : fallbackReason) + " No signature diagnostics were returned.";
	}

	private static void RecordDeepMetadataSignatureFailure(StandardLoadableFamilyDeepMetadata metadata, Family parentFamily, string fingerprintDebugRunFolder, string fingerprintDebugContextName, string reason)
	{
		if (metadata != null && parentFamily != null && string.IsNullOrWhiteSpace(metadata.ContentSignatureDebugPath))
		{
			LoadableFamilyContentSignatureResult signatureResult = LoadableFamilyContentSignatureService.BuildDiagnosticFailureResult("Precise", reason, parentFamily);
			metadata.ContentFingerprint = string.Empty;
			metadata.ContentFingerprintFailureReason = reason ?? string.Empty;
			metadata.ContentSignatureDebugPath = FingerprintDebugSignatureStore.SaveLoadableSignature(fingerprintDebugRunFolder, "standard", fingerprintDebugContextName, FamilyBrowserFamilyClassificationService.ResolveCategoryName(parentFamily), ((Element)parentFamily).Name ?? string.Empty, metadata.ContentFingerprint, signatureResult);
		}
	}

	private static void ApplyDialogCancelFingerprintFailure(StandardLoadableFamilySnapshotItem item, Family family, IEnumerable<FamilyThumbnailAutoConfirmedDialogRecord> records, string fingerprintDebugRunFolder, string fingerprintDebugContextName)
	{
		if (item != null && family != null && string.IsNullOrWhiteSpace(item.ContentFingerprint))
		{
			string reason = FamilyThumbnailConstraintDialogGuard.BuildFingerprintCanceledReason(records);
			if (!string.IsNullOrWhiteSpace(reason))
			{
				item.ContentFingerprintFailureReason = reason;
				LoadableFamilyContentSignatureResult signatureResult = LoadableFamilyContentSignatureService.BuildDiagnosticFailureResult("Precise", reason, family);
				item.ContentSignatureDebugPath = FingerprintDebugSignatureStore.SaveLoadableSignature(fingerprintDebugRunFolder, "standard", fingerprintDebugContextName, item.CategoryName, item.FamilyName, string.Empty, signatureResult);
			}
		}
	}

	private static void EnsureStandardLoadableFingerprints(StandardLibrarySnapshot snapshot, string operationName)
	{
		if (snapshot == null)
		{
			return;
		}
		List<StandardLoadableFamilySnapshotItem> failed = (snapshot.LoadableFamilies ?? new List<StandardLoadableFamilySnapshotItem>()).Where([SpecialName] (StandardLoadableFamilySnapshotItem x) => x != null && string.IsNullOrWhiteSpace(x.ContentFingerprint)).ToList();
		if (failed.Count != 0)
		{
			string details = string.Join(Environment.NewLine, failed.Select([SpecialName] (StandardLoadableFamilySnapshotItem x) => "- " + (x.CategoryName ?? string.Empty) + " / " + (x.FamilyName ?? string.Empty) + " / reason=" + (string.IsNullOrWhiteSpace(x.ContentFingerprintFailureReason) ? "Fingerprint was empty without a recorded reason." : x.ContentFingerprintFailureReason) + (string.IsNullOrWhiteSpace(x.ContentSignatureDebugPath) ? string.Empty : (" / " + x.ContentSignatureDebugPath))));
			throw new InvalidOperationException("STANDARD_RVT_FINGERPRINT_FAILED: " + (operationName ?? "Standard RVT scan") + " failed because one or more standard loadable family fingerprints were not created." + Environment.NewLine + "The standard snapshot was not accepted. Re-run the scan after fixing the listed family/families." + Environment.NewLine + details);
		}
	}

	private static void EnsureStandardLoadableFingerprint(StandardLoadableFamilySnapshotItem item, string operationName)
	{
		if (item != null && string.IsNullOrWhiteSpace(item.ContentFingerprint))
		{
			throw new InvalidOperationException("STANDARD_RVT_FINGERPRINT_FAILED: " + (operationName ?? "Standard RVT scan") + " failed because a standard loadable family fingerprint was not created." + Environment.NewLine + "The standard snapshot was not accepted. Re-run the scan after fixing the listed family." + Environment.NewLine + "- " + (item.CategoryName ?? string.Empty) + " / " + (item.FamilyName ?? string.Empty) + " / reason=" + (string.IsNullOrWhiteSpace(item.ContentFingerprintFailureReason) ? "Fingerprint was empty without a recorded reason." : item.ContentFingerprintFailureReason) + (string.IsNullOrWhiteSpace(item.ContentSignatureDebugPath) ? string.Empty : (" / " + item.ContentSignatureDebugPath)));
		}
	}

	private static string BuildPreciseScanThumbnailPath(string workspaceRoot, string sourceId, string categoryName, string familyName)
	{
		string BuildPreciseScanThumbnailPath;
		if (string.IsNullOrWhiteSpace(sourceId) || string.IsNullOrWhiteSpace(familyName))
		{
			BuildPreciseScanThumbnailPath = string.Empty;
		}
		else
		{
			try
			{
				BuildPreciseScanThumbnailPath = FamilyThumbnailPreviewService.GetCachedImagePath(workspaceRoot, sourceId, categoryName, familyName);
			}
			catch (Exception projectError)
			{
				ProjectData.SetProjectError(projectError);
				BuildPreciseScanThumbnailPath = string.Empty;
				ProjectData.ClearProjectError();
			}
		}
		return BuildPreciseScanThumbnailPath;
	}

	private static bool TryGeneratePreciseScanThumbnail(Document familyDoc, string previewImagePath)
	{
		bool TryGeneratePreciseScanThumbnail;
		if (familyDoc == null || string.IsNullOrWhiteSpace(previewImagePath))
		{
			TryGeneratePreciseScanThumbnail = false;
		}
		else
		{
			try
			{
				FamilyThumbnailPreviewService.GenerateFromOpenFamilyDocument(familyDoc, previewImagePath);
				TryGeneratePreciseScanThumbnail = true;
			}
			catch (Exception projectError)
			{
				ProjectData.SetProjectError(projectError);
				TryGeneratePreciseScanThumbnail = false;
				ProjectData.ClearProjectError();
			}
		}
		return TryGeneratePreciseScanThumbnail;
	}

	private static ISet<string> NormalizeFamilyNameSet(ISet<string> values)
	{
		HashSet<string> result = new HashSet<string>(StringComparer.Ordinal);
		if (values == null)
		{
			return result;
		}
		foreach (string value in values)
		{
			string normalized = Normalize(value);
			if (!string.IsNullOrWhiteSpace(normalized))
			{
				result.Add(normalized);
			}
		}
		return result;
	}

	private static List<StandardFamilyParameterSnapshotItem> CaptureFamilyDocumentParameters(Document familyDoc)
	{
		return FamilyDocumentParameterCaptureService.Capture(familyDoc);
	}

	private static bool ShouldCaptureFamilyManagerParameter(FamilyParameter familyParameter)
	{
		if (familyParameter == null || familyParameter.Definition == null)
		{
			return false;
		}
		if (string.IsNullOrWhiteSpace(ResolveFamilyParameterName(familyParameter)))
		{
			return false;
		}
		int idValue = ResolveFamilyParameterIdInteger(familyParameter);
		if (idValue == -1002001)
		{
			return false;
		}
		if (idValue < 0 && !IsSharedFamilyParameter(familyParameter))
		{
			return false;
		}
		return true;
	}

	private static string ResolveCurrentFamilyTypeName(FamilyManager manager)
	{
		try
		{
			if (manager != null && manager.CurrentType != null)
			{
				return manager.CurrentType.Name ?? string.Empty;
			}
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ProjectData.ClearProjectError();
		}
		return string.Empty;
	}

	private static string ResolveFamilyParameterName(FamilyParameter familyParameter)
	{
		string ResolveFamilyParameterName;
		try
		{
			object obj;
			if (familyParameter == null)
			{
				obj = null;
			}
			else
			{
				Definition definition = familyParameter.Definition;
				obj = ((definition != null) ? definition.Name : null);
			}
			if (obj == null)
			{
				obj = string.Empty;
			}
			ResolveFamilyParameterName = (string)obj;
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ResolveFamilyParameterName = string.Empty;
			ProjectData.ClearProjectError();
		}
		return ResolveFamilyParameterName;
	}

	private static string ResolveFamilyParameterStorageTypeName(FamilyParameter familyParameter)
	{
		//IL_0001: Unknown result type (might be due to invalid IL or missing references)
		//IL_0006: Unknown result type (might be due to invalid IL or missing references)
		string ResolveFamilyParameterStorageTypeName;
		try
		{
			ResolveFamilyParameterStorageTypeName = ((Enum)familyParameter.StorageType/*cast due to .constrained prefix*/).ToString();
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ResolveFamilyParameterStorageTypeName = string.Empty;
			ProjectData.ClearProjectError();
		}
		return ResolveFamilyParameterStorageTypeName;
	}

	private static string ResolveFamilyParameterId(FamilyParameter familyParameter)
	{
		int idValue = ResolveFamilyParameterIdInteger(familyParameter);
		if (idValue == int.MinValue)
		{
			return string.Empty;
		}
		return idValue.ToString(CultureInfo.InvariantCulture);
	}

	private static int ResolveFamilyParameterIdInteger(FamilyParameter familyParameter)
	{
		try
		{
			if (familyParameter != null && familyParameter.Id != null)
			{
				return RevitElementIdCompat.CompatIntegerValue(familyParameter.Id);
			}
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ProjectData.ClearProjectError();
		}
		return int.MinValue;
	}

	private static bool IsSharedFamilyParameter(FamilyParameter familyParameter)
	{
		try
		{
			if (familyParameter != null && familyParameter.IsShared)
			{
				return true;
			}
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ProjectData.ClearProjectError();
		}
		return !string.IsNullOrWhiteSpace(ResolveFamilyParameterExternalGuid(familyParameter));
	}

	private static string ResolveFamilyParameterExternalGuid(FamilyParameter familyParameter)
	{
		try
		{
			Definition obj = ((familyParameter != null) ? familyParameter.Definition : null);
			ExternalDefinition externalDefinition = (ExternalDefinition)(object)((obj is ExternalDefinition) ? obj : null);
			if (externalDefinition != null)
			{
				return externalDefinition.GUID.ToString("D");
			}
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ProjectData.ClearProjectError();
		}
		return string.Empty;
	}

	private static string ResolveFamilyParameterFormula(FamilyParameter familyParameter)
	{
		string ResolveFamilyParameterFormula;
		if (familyParameter == null)
		{
			ResolveFamilyParameterFormula = string.Empty;
		}
		else
		{
			try
			{
				PropertyInfo propertyInfo = ((object)familyParameter).GetType().GetProperty("Formula");
				ResolveFamilyParameterFormula = (((object)propertyInfo != null) ? NormalizeMultiline(Convert.ToString(RuntimeHelpers.GetObjectValue(RuntimeHelpers.GetObjectValue(propertyInfo.GetValue(familyParameter, null))), CultureInfo.InvariantCulture)) : string.Empty);
			}
			catch (Exception projectError)
			{
				ProjectData.SetProjectError(projectError);
				ResolveFamilyParameterFormula = string.Empty;
				ProjectData.ClearProjectError();
			}
		}
		return ResolveFamilyParameterFormula;
	}

	private static string ResolveFamilyParameterValue(Document familyDoc, FamilyManager manager, FamilyParameter familyParameter)
	{
		//IL_0021: Unknown result type (might be due to invalid IL or missing references)
		//IL_0026: Unknown result type (might be due to invalid IL or missing references)
		//IL_0027: Unknown result type (might be due to invalid IL or missing references)
		//IL_0029: Unknown result type (might be due to invalid IL or missing references)
		//IL_003f: Expected I4, but got Unknown
		string ResolveFamilyParameterValue;
		try
		{
			if (manager == null || manager.CurrentType == null || familyParameter == null)
			{
				ResolveFamilyParameterValue = string.Empty;
			}
			else
			{
				FamilyType familyType = manager.CurrentType;
				StorageType storageType = familyParameter.StorageType;
				switch (storageType - 1)
				{
				case 2:
					ResolveFamilyParameterValue = NormalizeMultiline(familyType.AsString(familyParameter));
					break;
				case 1:
				{
					object valueObject = familyType.AsDouble(familyParameter);
					ResolveFamilyParameterValue = ((valueObject != null) ? Convert.ToDouble(RuntimeHelpers.GetObjectValue(valueObject), CultureInfo.InvariantCulture).ToString("G17", CultureInfo.InvariantCulture) : string.Empty);
					break;
				}
				case 0:
				{
					object valueObject2 = familyType.AsInteger(familyParameter);
					ResolveFamilyParameterValue = ((valueObject2 != null) ? Convert.ToInt32(RuntimeHelpers.GetObjectValue(valueObject2), CultureInfo.InvariantCulture).ToString(CultureInfo.InvariantCulture) : string.Empty);
					break;
				}
				case 3:
				{
					ElementId id = familyType.AsElementId(familyParameter);
					if (id == null || id == ElementId.InvalidElementId)
					{
						ResolveFamilyParameterValue = string.Empty;
						break;
					}
					Element referenced = ((familyDoc == null) ? null : familyDoc.GetElement(id));
					ResolveFamilyParameterValue = ((referenced == null) ? RevitElementIdCompat.CompatIntegerValue(id).ToString(CultureInfo.InvariantCulture) : (((object)referenced).GetType().Name + ":" + ResolveElementName(referenced)));
					break;
				}
				default:
					ResolveFamilyParameterValue = string.Empty;
					break;
				}
			}
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ResolveFamilyParameterValue = string.Empty;
			ProjectData.ClearProjectError();
		}
		return ResolveFamilyParameterValue;
	}

	private static string WriteRegistrationDialogDiagnosticReport(string workspaceRoot, string sourceId, IEnumerable<FamilyThumbnailAutoConfirmedDialogRecord> records)
	{
		List<FamilyThumbnailAutoConfirmedDialogRecord> list = (records ?? Enumerable.Empty<FamilyThumbnailAutoConfirmedDialogRecord>()).Where([SpecialName] (FamilyThumbnailAutoConfirmedDialogRecord x) => x != null).ToList();
		string WriteRegistrationDialogDiagnosticReport;
		if (list.Count == 0)
		{
			WriteRegistrationDialogDiagnosticReport = string.Empty;
		}
		else
		{
			try
			{
				FamilyBrowserStandardPolicyStore.RequireManagedDataRootForWrite(workspaceRoot, T("Save standard scan diagnostics", "표준 스캔 진단 저장"));
				string snapshotFolder = FamilyBrowserStandardPolicyStore.GetSnapshotFolder(workspaceRoot);
				Directory.CreateDirectory(snapshotFolder);
				string reportPath = Path.Combine(snapshotFolder, "standard-scan-dialog-diagnostics-" + SafeFileNameToken(sourceId ?? "standard") + "-" + DateTime.UtcNow.ToString("yyyyMMdd-HHmmss", CultureInfo.InvariantCulture) + ".txt");
				StringBuilder sb = new StringBuilder();
				sb.AppendLine("KKY Family Browser Standard RVT Scan Dialog Diagnostics");
				sb.AppendLine("Created UTC: " + DateTime.UtcNow.ToString("u", CultureInfo.InvariantCulture));
				sb.AppendLine("Policy: Remove Constraints/Delete Constraints dialogs are cancelled so standard families are not modified. Geometry/constraint warnings that can continue are confirmed with OK and the scan continues.");
				sb.AppendLine("Auto-handled dialogs: " + list.Count.ToString(CultureInfo.InvariantCulture));
				sb.AppendLine();
				foreach (FamilyThumbnailAutoConfirmedDialogRecord record in list)
				{
					sb.AppendLine("- " + (string.IsNullOrWhiteSpace(record.CategoryName) ? "-" : record.CategoryName) + " | " + (string.IsNullOrWhiteSpace(record.FamilyName) ? "-" : record.FamilyName) + " | " + (string.IsNullOrWhiteSpace(record.Reason) ? "-" : record.Reason) + " | result=" + (string.IsNullOrWhiteSpace(record.OverrideResult) ? "-" : record.OverrideResult) + " | utc=" + (string.IsNullOrWhiteSpace(record.ConfirmedAtUtc) ? "-" : record.ConfirmedAtUtc));
					if (!string.IsNullOrWhiteSpace(record.DialogText))
					{
						sb.AppendLine("  Dialog: " + CompactDiagnosticText(record.DialogText));
					}
				}
				File.WriteAllText(reportPath, sb.ToString(), Encoding.UTF8);
				WriteRegistrationDialogDiagnosticReport = reportPath;
			}
			catch (Exception projectError)
			{
				ProjectData.SetProjectError(projectError);
				WriteRegistrationDialogDiagnosticReport = string.Empty;
				ProjectData.ClearProjectError();
			}
		}
		return WriteRegistrationDialogDiagnosticReport;
	}

	private static void TryAddNestedLoadableFamily(IDictionary<string, StandardNestedLoadableFamilySnapshotItem> result, Document familyDoc, string parentKey, Family nestedFamily)
	{
		if (result == null || familyDoc == null || ShouldIgnoreNestedFamilyReference(nestedFamily) || !ResolveIsShared(nestedFamily))
		{
			return;
		}
		string categoryName = FamilyBrowserFamilyClassificationService.ResolveCategoryName(nestedFamily);
		string categoryId = FamilyBrowserFamilyClassificationService.ResolveCategoryId(nestedFamily);
		string categoryGroup = FamilyBrowserFamilyClassificationService.ResolveCategoryGroup(nestedFamily);
		if (!string.Equals(categoryGroup, "Model", StringComparison.OrdinalIgnoreCase))
		{
			return;
		}
		string familyName = ((Element)nestedFamily).Name ?? string.Empty;
		string nestedKey = BuildNestedLoadableKey(categoryId, categoryName, familyName);
		if (!string.IsNullOrWhiteSpace(nestedKey) && !string.Equals(nestedKey, parentKey, StringComparison.Ordinal) && !result.ContainsKey(nestedKey))
		{
			List<string> typeNames = (from x in nestedFamily.GetFamilySymbolIds().Select([SpecialName] (ElementId id) =>
				{
					Element element = familyDoc.GetElement(id);
					return (ElementType)(object)((element is ElementType) ? element : null);
				})
				where x != null
				select ((Element)x).Name ?? string.Empty).OrderBy<string, string>([SpecialName] (string x) => Normalize(x), StringComparer.Ordinal).ToList();
			result.Add(nestedKey, new StandardNestedLoadableFamilySnapshotItem
			{
				FamilyName = familyName,
				CategoryName = categoryName,
				CategoryId = categoryId,
				CategoryGroup = categoryGroup,
				TypeCount = typeNames.Count,
				TypeNames = typeNames,
				IsShared = ResolveIsShared(nestedFamily)
			});
		}
	}

	private static Family ResolveNestedFamily(FamilyInstance instance)
	{
		Family ResolveNestedFamily;
		try
		{
			FamilySymbol symbol = ((instance != null) ? instance.Symbol : null);
			ResolveNestedFamily = ((symbol != null) ? symbol.Family : null);
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ResolveNestedFamily = null;
			ProjectData.ClearProjectError();
		}
		return ResolveNestedFamily;
	}

	private static string SafeFileNameToken(string value)
	{
		string text = value ?? string.Empty;
		if (string.IsNullOrWhiteSpace(text))
		{
			return "standard";
		}
		char[] invalidFileNameChars = Path.GetInvalidFileNameChars();
		string safe = new string(text.Select([SpecialName] (char ch) => (!Enumerable.Contains(invalidFileNameChars, ch)) ? ch : '_').ToArray()).Trim().TrimEnd(new char[2] { '.', ' ' });
		if (string.IsNullOrWhiteSpace(safe))
		{
			return "standard";
		}
		if (safe.Length > 80)
		{
			safe = safe.Substring(0, 80);
		}
		return safe;
	}

	private static string CompactDiagnosticText(string value)
	{
		string text = (value ?? string.Empty).Replace("\r", " ").Replace("\n", " ").Trim();
		while (text.Contains("  "))
		{
			text = text.Replace("  ", " ");
		}
		if (text.Length > 900)
		{
			return text.Substring(0, 900) + "...";
		}
		return text;
	}

	private static bool ShouldIgnoreNestedFamilyReference(Family family)
	{
		if (family == null)
		{
			return true;
		}
		try
		{
			if (family.IsInPlace)
			{
				return true;
			}
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ProjectData.ClearProjectError();
		}
		return FamilyBrowserFamilyClassificationService.IsTypeManagedFamilyLike(FamilyBrowserFamilyClassificationService.ResolveCategoryName(family), FamilyBrowserFamilyClassificationService.ResolveCategoryId(family), ((Element)family).Name ?? string.Empty);
	}

	private static void MarkNestedLoadableChildren(StandardLibrarySnapshot snapshot)
	{
		if (snapshot == null || snapshot.LoadableFamilies == null)
		{
			return;
		}
		HashSet<string> nestedKeys = new HashSet<string>(StringComparer.Ordinal);
		HashSet<string> nestedFamilyNames = new HashSet<string>(StringComparer.Ordinal);
		foreach (StandardLoadableFamilySnapshotItem parentItem in snapshot.LoadableFamilies)
		{
			if (parentItem == null || parentItem.NestedLoadableFamilies == null)
			{
				continue;
			}
			foreach (StandardNestedLoadableFamilySnapshotItem child in parentItem.NestedLoadableFamilies)
			{
				if (child != null && child.IsShared && IsModelNestedLoadableChild(child))
				{
					string key = BuildNestedLoadableKey(child);
					if (!string.IsNullOrWhiteSpace(key))
					{
						nestedKeys.Add(key);
					}
					string familyNameKey = Normalize((child == null) ? string.Empty : child.FamilyName);
					if (!string.IsNullOrWhiteSpace(familyNameKey))
					{
						nestedFamilyNames.Add(familyNameKey);
					}
				}
			}
		}
		foreach (StandardLoadableFamilySnapshotItem item in snapshot.LoadableFamilies)
		{
			if (item != null)
			{
				item.IsNestedLoadableChild = item.IsShared && IsModelLoadableFamily(item) && (nestedKeys.Contains(BuildNestedLoadableKey(item)) || nestedFamilyNames.Contains(Normalize(item.FamilyName)));
			}
		}
	}

	private static string BuildNestedLoadableKey(StandardLoadableFamilySnapshotItem item)
	{
		if (item == null)
		{
			return string.Empty;
		}
		return StandardLibraryRegistrationService.BuildNestedLoadableKey(item.CategoryId, item.CategoryName, item.FamilyName);
	}

	private static string BuildNestedLoadableKey(StandardNestedLoadableFamilySnapshotItem item)
	{
		if (item == null)
		{
			return string.Empty;
		}
		return StandardLibraryRegistrationService.BuildNestedLoadableKey(item.CategoryId, item.CategoryName, item.FamilyName);
	}

	private static bool IsModelNestedLoadableChild(StandardNestedLoadableFamilySnapshotItem item)
	{
		if (item == null)
		{
			return false;
		}
		return string.Equals(FamilyBrowserFamilyClassificationService.ResolveCategoryGroup(item.CategoryGroup, item.CategoryName, item.CategoryId, item.FamilyName), "Model", StringComparison.OrdinalIgnoreCase);
	}

	private static bool IsModelLoadableFamily(StandardLoadableFamilySnapshotItem item)
	{
		if (item == null)
		{
			return false;
		}
		return string.Equals(FamilyBrowserFamilyClassificationService.ResolveCategoryGroup(item.CategoryGroup, item.CategoryName, item.CategoryId, item.FamilyName), "Model", StringComparison.OrdinalIgnoreCase);
	}

	private static string BuildNestedLoadableKey(string categoryId, string categoryName, string familyName)
	{
		string categoryKey = Normalize(categoryId);
		if (categoryKey.Length == 0)
		{
			categoryKey = Normalize(categoryName);
		}
		string familyKey = Normalize(familyName);
		if (categoryKey.Length == 0 || familyKey.Length == 0)
		{
			return string.Empty;
		}
		return categoryKey + "|" + familyKey;
	}

	private static bool IsBrowserLoadableFamily(Family family)
	{
		if (family == null)
		{
			return false;
		}
		try
		{
			if (family.IsInPlace)
			{
				return false;
			}
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ProjectData.ClearProjectError();
		}
		return !IsTypeManagedFamilyCategory(ResolveCategoryName(family), ((Element)family).Name ?? string.Empty);
	}

	private static bool IsTypeManagedFamilyCategory(string categoryName, string familyName)
	{
		string compact = Normalize(categoryName + " " + familyName).Replace(" ", string.Empty);
		if (string.IsNullOrWhiteSpace(compact))
		{
			return false;
		}
		if (compact.Contains("mullion") || compact.Contains("멀리언"))
		{
			return true;
		}
		return false;
	}

	private static Document FindOpenDocument(Application application, string resolvedPath)
	{
		//IL_0014: Unknown result type (might be due to invalid IL or missing references)
		//IL_001a: Expected O, but got Unknown
		foreach (Document document in application.Documents)
		{
			Document doc = document;
			if (doc != null && !string.IsNullOrWhiteSpace(doc.PathName))
			{
				string candidatePath = string.Empty;
				try
				{
					candidatePath = Path.GetFullPath(doc.PathName);
				}
				catch (Exception projectError)
				{
					ProjectData.SetProjectError(projectError);
					ProjectData.ClearProjectError();
					continue;
				}
				if (string.Equals(candidatePath, resolvedPath, StringComparison.OrdinalIgnoreCase))
				{
					return doc;
				}
			}
		}
		return null;
	}

	private static string InferSourceKind(string resolvedPath)
	{
		if (resolvedPath.StartsWith("\\\\", StringComparison.OrdinalIgnoreCase))
		{
			return "NetworkShareRvt";
		}
		return "LocalFileRvt";
	}

	private static string NormalizeSnapshotMode(string value)
	{
		if (string.Equals(value, "Precise", StringComparison.OrdinalIgnoreCase))
		{
			return "Precise";
		}
		return "Fast";
	}

	private static string BuildSourceId(string sourceKind, string resolvedPath)
	{
		string normalized = sourceKind + "|" + Normalize(resolvedPath);
		using SHA256 sha = SHA256.Create();
		byte[] bytes = Encoding.UTF8.GetBytes(normalized);
		return BitConverter.ToString(sha.ComputeHash(bytes)).Replace("-", string.Empty).Substring(0, 16);
	}

	private static string ResolveCategoryName(Element element)
	{
		string ResolveCategoryName;
		try
		{
			Category category = element.Category;
			ResolveCategoryName = ((category != null) ? category.Name : null) ?? string.Empty;
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ResolveCategoryName = string.Empty;
			ProjectData.ClearProjectError();
		}
		return ResolveCategoryName;
	}

	private static string ResolveCategoryId(Element element)
	{
		string ResolveCategoryId;
		try
		{
			ResolveCategoryId = ((element != null && element.Category != null && element.Category.Id != null) ? RevitElementIdCompat.CompatIntegerValue(element.Category.Id).ToString(CultureInfo.InvariantCulture) : string.Empty);
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ResolveCategoryId = string.Empty;
			ProjectData.ClearProjectError();
		}
		return ResolveCategoryId;
	}

	private static string ResolveCategoryName(Family family)
	{
		string ResolveCategoryName;
		try
		{
			Category familyCategory = family.FamilyCategory;
			ResolveCategoryName = ((familyCategory != null) ? familyCategory.Name : null) ?? string.Empty;
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ResolveCategoryName = string.Empty;
			ProjectData.ClearProjectError();
		}
		return ResolveCategoryName;
	}

	private static string ResolveCategoryId(Family family)
	{
		string ResolveCategoryId;
		try
		{
			ResolveCategoryId = ((family != null && family.FamilyCategory != null && family.FamilyCategory.Id != null) ? RevitElementIdCompat.CompatIntegerValue(family.FamilyCategory.Id).ToString(CultureInfo.InvariantCulture) : string.Empty);
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ResolveCategoryId = string.Empty;
			ProjectData.ClearProjectError();
		}
		return ResolveCategoryId;
	}

	private static string ResolveFamilyCategoryGroup(Family family)
	{
		//IL_0019: Unknown result type (might be due to invalid IL or missing references)
		//IL_001e: Unknown result type (might be due to invalid IL or missing references)
		//IL_001f: Unknown result type (might be due to invalid IL or missing references)
		//IL_0021: Invalid comparison between Unknown and I4
		//IL_0023: Unknown result type (might be due to invalid IL or missing references)
		//IL_0025: Invalid comparison between Unknown and I4
		string ResolveFamilyCategoryGroup;
		try
		{
			if (family == null || family.FamilyCategory == null)
			{
				ResolveFamilyCategoryGroup = string.Empty;
			}
			else
			{
				CategoryType categoryType = family.FamilyCategory.CategoryType;
				ResolveFamilyCategoryGroup = (((int)categoryType == 1) ? "Model" : (((int)categoryType != 2) ? "Other" : "Annotation"));
			}
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ResolveFamilyCategoryGroup = string.Empty;
			ProjectData.ClearProjectError();
		}
		return ResolveFamilyCategoryGroup;
	}

	private static bool ResolveIsShared(Family family)
	{
		//IL_0011: Unknown result type (might be due to invalid IL or missing references)
		//IL_0017: Invalid comparison between Unknown and I4
		try
		{
			Parameter parameterValue = ((Element)family)[(BuiltInParameter)(-1012834)];
			if (parameterValue != null && (int)parameterValue.StorageType == 1)
			{
				return parameterValue.AsInteger() != 0;
			}
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ProjectData.ClearProjectError();
		}
		return false;
	}

	private static List<StandardFamilyParameterSnapshotItem> CaptureLoadableFamilyParameters(Document doc, Family family)
	{
		List<StandardFamilyParameterSnapshotItem> result = new List<StandardFamilyParameterSnapshotItem>();
		if (doc == null || family == null)
		{
			return result;
		}
		foreach (Parameter parameter in (from Parameter x in (IEnumerable)((Element)family).Parameters
			where ShouldCaptureParameter(x)
			select x).OrderBy<Parameter, string>([SpecialName] (Parameter x) => Normalize(ResolveParameterName(x)), StringComparer.Ordinal))
		{
			result.Add(BuildParameterSnapshot(doc, parameter, "Family", string.Empty));
		}
		foreach (ElementId symbolId in family.GetFamilySymbolIds())
		{
			Element element = doc.GetElement(symbolId);
			FamilySymbol symbol = (FamilySymbol)(object)((element is FamilySymbol) ? element : null);
			if (symbol == null)
			{
				continue;
			}
			string typeName = ((Element)symbol).Name ?? string.Empty;
			foreach (Parameter parameter2 in (from Parameter x in (IEnumerable)((Element)symbol).Parameters
				where ShouldCaptureParameter(x)
				select x).OrderBy<Parameter, string>([SpecialName] (Parameter x) => Normalize(ResolveParameterName(x)), StringComparer.Ordinal))
			{
				result.Add(BuildParameterSnapshot(doc, parameter2, "Type", typeName));
			}
		}
		return DeduplicateParameterSnapshots(result);
	}

	private static List<StandardFamilyParameterSnapshotItem> DeduplicateParameterSnapshots(IEnumerable<StandardFamilyParameterSnapshotItem> parameters)
	{
		return FamilyParameterSnapshotNormalizationService.DeduplicateDefinitions(parameters);
	}

	private static string BuildParameterSnapshotIdentityKey(StandardFamilyParameterSnapshotItem item)
	{
		if (item == null)
		{
			return string.Empty;
		}
		string guid = Normalize(item.ExternalGuid);
		if (guid.Length > 0)
		{
			return "guid|" + guid;
		}
		if (item.IsShared)
		{
			return "shared|" + Normalize(item.Name) + "|" + Normalize(item.StorageType) + "|" + Normalize(item.Formula);
		}
		return "local|" + Normalize(item.Scope) + "|" + Normalize(item.Name) + "|" + Normalize(item.StorageType) + "|" + item.IsInstance + "|" + Normalize(item.Formula) + "|" + Normalize(item.ParameterId);
	}

	private static StandardFamilyParameterSnapshotItem SelectPreferredParameterSnapshot(IEnumerable<StandardFamilyParameterSnapshotItem> items)
	{
		return (from x in items ?? Enumerable.Empty<StandardFamilyParameterSnapshotItem>()
			where x != null
			orderby ParameterSnapshotPreferenceRank(x), !string.IsNullOrWhiteSpace(x.Formula) descending, !string.IsNullOrWhiteSpace(x.ValuePreview) descending, !string.IsNullOrWhiteSpace(x.TypeName) descending
			select x).FirstOrDefault();
	}

	private static int ParameterSnapshotPreferenceRank(StandardFamilyParameterSnapshotItem item)
	{
		if (item == null)
		{
			return int.MaxValue;
		}
		string scope = Normalize(item.Scope);
		if (item.IsInstance || string.Equals(scope, "instance", StringComparison.Ordinal))
		{
			return 0;
		}
		if (string.Equals(scope, "type", StringComparison.Ordinal))
		{
			return 1;
		}
		if (string.Equals(scope, "family", StringComparison.Ordinal))
		{
			return 2;
		}
		return 3;
	}

	private static StandardFamilyParameterSnapshotItem BuildParameterSnapshot(Document doc, Parameter parameter, string scope, string typeName)
	{
		return new StandardFamilyParameterSnapshotItem
		{
			Scope = (scope ?? string.Empty),
			TypeName = (typeName ?? string.Empty),
			Name = ResolveParameterName(parameter),
			StorageType = ResolveStorageTypeName(parameter),
			ValuePreview = ResolveParameterValue(doc, parameter),
			Formula = string.Empty,
			IsInstance = false,
			IsReadOnly = SafeBool([SpecialName] () => parameter.IsReadOnly),
			IsShared = SafeBool([SpecialName] () => parameter.IsShared),
			ParameterId = ResolveParameterId(parameter),
			ExternalGuid = ResolveExternalGuid(parameter)
		};
	}

	private static bool ShouldCaptureParameter(Parameter parameter)
	{
		if (parameter == null || parameter.Definition == null)
		{
			return false;
		}
		if (string.IsNullOrWhiteSpace(ResolveParameterName(parameter)))
		{
			return false;
		}
		int idValue = ResolveParameterIdInteger(parameter);
		if (idValue == -1002001)
		{
			return false;
		}
		if (idValue < 0 && !SafeBool([SpecialName] () => parameter.IsShared))
		{
			return false;
		}
		return true;
	}

	private static string ResolveParameterName(Parameter parameter)
	{
		string ResolveParameterName;
		try
		{
			object obj;
			if (parameter == null)
			{
				obj = null;
			}
			else
			{
				Definition definition = parameter.Definition;
				obj = ((definition != null) ? definition.Name : null);
			}
			if (obj == null)
			{
				obj = string.Empty;
			}
			ResolveParameterName = (string)obj;
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ResolveParameterName = string.Empty;
			ProjectData.ClearProjectError();
		}
		return ResolveParameterName;
	}

	private static string ResolveStorageTypeName(Parameter parameter)
	{
		//IL_0001: Unknown result type (might be due to invalid IL or missing references)
		//IL_0006: Unknown result type (might be due to invalid IL or missing references)
		string ResolveStorageTypeName;
		try
		{
			ResolveStorageTypeName = ((Enum)parameter.StorageType/*cast due to .constrained prefix*/).ToString();
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ResolveStorageTypeName = string.Empty;
			ProjectData.ClearProjectError();
		}
		return ResolveStorageTypeName;
	}

	private static string ResolveParameterId(Parameter parameter)
	{
		int idValue = ResolveParameterIdInteger(parameter);
		if (idValue != int.MinValue)
		{
			return idValue.ToString(CultureInfo.InvariantCulture);
		}
		return string.Empty;
	}

	private static int ResolveParameterIdInteger(Parameter parameter)
	{
		try
		{
			if (parameter != null && parameter.Id != null)
			{
				return RevitElementIdCompat.CompatIntegerValue(parameter.Id);
			}
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ProjectData.ClearProjectError();
		}
		return int.MinValue;
	}

	private static string ResolveExternalGuid(Parameter parameter)
	{
		try
		{
			Definition obj = ((parameter != null) ? parameter.Definition : null);
			ExternalDefinition externalDefinition = (ExternalDefinition)(object)((obj is ExternalDefinition) ? obj : null);
			if (externalDefinition != null)
			{
				return externalDefinition.GUID.ToString("D");
			}
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ProjectData.ClearProjectError();
		}
		return string.Empty;
	}

	private static string ResolveParameterValue(Document doc, Parameter parameter)
	{
		//IL_0001: Unknown result type (might be due to invalid IL or missing references)
		//IL_0006: Unknown result type (might be due to invalid IL or missing references)
		//IL_0007: Unknown result type (might be due to invalid IL or missing references)
		//IL_0009: Unknown result type (might be due to invalid IL or missing references)
		//IL_001f: Expected I4, but got Unknown
		string ResolveParameterValue;
		try
		{
			StorageType storageType = parameter.StorageType;
			switch (storageType - 1)
			{
			case 2:
				ResolveParameterValue = NormalizeMultiline(parameter.AsString());
				break;
			case 1:
			{
				string formatted2 = parameter.AsValueString();
				ResolveParameterValue = (string.IsNullOrWhiteSpace(formatted2) ? parameter.AsDouble().ToString("G17", CultureInfo.InvariantCulture) : NormalizeMultiline(formatted2));
				break;
			}
			case 0:
			{
				string formatted3 = parameter.AsValueString();
				ResolveParameterValue = (string.IsNullOrWhiteSpace(formatted3) ? parameter.AsInteger().ToString(CultureInfo.InvariantCulture) : NormalizeMultiline(formatted3));
				break;
			}
			case 3:
			{
				ElementId id = parameter.AsElementId();
				if (id == null || id == ElementId.InvalidElementId)
				{
					ResolveParameterValue = string.Empty;
				}
				else if (RevitElementIdCompat.CompatIntegerValue(id) < 0)
				{
					string formatted = parameter.AsValueString();
					ResolveParameterValue = (string.IsNullOrWhiteSpace(formatted) ? RevitElementIdCompat.CompatIntegerValue(id).ToString(CultureInfo.InvariantCulture) : NormalizeMultiline(formatted));
				}
				else
				{
					Element element = ((doc == null) ? null : doc.GetElement(id));
					ResolveParameterValue = ((element != null) ? BuildReferencedElementParameterValue(doc, element) : RevitElementIdCompat.CompatIntegerValue(id).ToString(CultureInfo.InvariantCulture));
				}
				break;
			}
			default:
				ResolveParameterValue = string.Empty;
				break;
			}
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ResolveParameterValue = string.Empty;
			ProjectData.ClearProjectError();
		}
		return ResolveParameterValue;
	}

	private static string BuildReferencedElementParameterValue(Document doc, Element element)
	{
		if (element == null)
		{
			return string.Empty;
		}
		List<string> parts = new List<string>
		{
			((object)element).GetType().Name,
			ResolveCategoryName(element),
			ResolveElementName(element)
		};
		if (!(element is FamilySymbol))
		{
			string signature = RoutingPartSignatureService.Build(doc, element);
			if (!string.IsNullOrWhiteSpace(signature))
			{
				parts.Add(signature);
			}
		}
		return NormalizeMultiline(string.Join("|", parts));
	}

	private static List<StandardSystemTypeLayerSnapshotItem> CaptureSystemTypeLayers(Document doc, ElementType systemType)
	{
		List<StandardSystemTypeLayerSnapshotItem> result = new List<StandardSystemTypeLayerSnapshotItem>();
		if (doc == null || systemType == null)
		{
			return result;
		}
		HostObjAttributes hostAttributes = (HostObjAttributes)(object)((systemType is HostObjAttributes) ? systemType : null);
		if (hostAttributes == null)
		{
			return result;
		}
		CompoundStructure compound = null;
		try
		{
			compound = hostAttributes.GetCompoundStructure();
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			compound = null;
			ProjectData.ClearProjectError();
		}
		if (compound == null)
		{
			return result;
		}
		try
		{
			int index = 1;
			foreach (CompoundStructureLayer layer in compound.GetLayers())
			{
				string functionName = SafeLayerFunctionName(layer);
				string materialName = ResolveLayerMaterialName(doc, layer);
				if (string.IsNullOrWhiteSpace(materialName))
				{
					materialName = functionName;
				}
				result.Add(new StandardSystemTypeLayerSnapshotItem
				{
					Index = index,
					FunctionName = functionName,
					MaterialName = materialName,
					ThicknessFeet = layer.Width,
					ThicknessDisplay = FormatLayerThickness(layer.Width)
				});
				index = checked(index + 1);
			}
		}
		catch (Exception projectError2)
		{
			ProjectData.SetProjectError(projectError2);
			ProjectData.ClearProjectError();
		}
		return result;
	}

	private static string SafeLayerFunctionName(CompoundStructureLayer layer)
	{
		//IL_0001: Unknown result type (might be due to invalid IL or missing references)
		//IL_0006: Unknown result type (might be due to invalid IL or missing references)
		string SafeLayerFunctionName;
		try
		{
			SafeLayerFunctionName = ((Enum)layer.Function/*cast due to .constrained prefix*/).ToString();
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			SafeLayerFunctionName = string.Empty;
			ProjectData.ClearProjectError();
		}
		return SafeLayerFunctionName;
	}

	private static string ResolveLayerMaterialName(Document doc, CompoundStructureLayer layer)
	{
		string ResolveLayerMaterialName;
		if (doc == null)
		{
			ResolveLayerMaterialName = string.Empty;
		}
		else
		{
			try
			{
				ResolveLayerMaterialName = ((layer.MaterialId != null && !(layer.MaterialId == ElementId.InvalidElementId)) ? ResolveElementName(doc.GetElement(layer.MaterialId)) : string.Empty);
			}
			catch (Exception projectError)
			{
				ProjectData.SetProjectError(projectError);
				ResolveLayerMaterialName = string.Empty;
				ProjectData.ClearProjectError();
			}
		}
		return ResolveLayerMaterialName;
	}

	private static string FormatLayerThickness(double widthFeet)
	{
		if (double.IsNaN(widthFeet) || double.IsInfinity(widthFeet))
		{
			return string.Empty;
		}
		double widthMillimeters = widthFeet * 304.8;
		if (Math.Abs(widthMillimeters) < 0.0005)
		{
			widthMillimeters = 0.0;
		}
		return widthMillimeters.ToString("0.###", CultureInfo.InvariantCulture) + " mm";
	}

	private static string ResolveElementName(Element element)
	{
		string ResolveElementName;
		if (element == null)
		{
			ResolveElementName = string.Empty;
		}
		else
		{
			try
			{
				ResolveElementName = element.Name ?? string.Empty;
			}
			catch (Exception projectError)
			{
				ProjectData.SetProjectError(projectError);
				ResolveElementName = string.Empty;
				ProjectData.ClearProjectError();
			}
		}
		return ResolveElementName;
	}

	private static string NormalizeMultiline(string value)
	{
		if (value == null)
		{
			return string.Empty;
		}
		return value.Replace("\r\n", " ").Replace("\r", " ").Replace("\n", " ")
			.Trim();
	}

	private static bool SafeBool(Func<bool> reader)
	{
		bool SafeBool;
		try
		{
			SafeBool = reader();
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			SafeBool = false;
			ProjectData.ClearProjectError();
		}
		return SafeBool;
	}

	private static string Normalize(string value)
	{
		return SystemTypeIdentityService.Normalize(value);
	}

	private static string ResolveSemanticFingerprint(IDictionary<string, string> semanticMap, string typeClassName, string categoryName, string typeName)
	{
		if (semanticMap == null)
		{
			return string.Empty;
		}
		string key = SystemTypeIdentityService.BuildKey(typeClassName, categoryName, typeName);
		string fingerprint = null;
		if (semanticMap.TryGetValue(key, out fingerprint))
		{
			return fingerprint ?? string.Empty;
		}
		return string.Empty;
	}

	private static SystemTypeSemanticSnapshot ResolveSemanticSnapshot(IDictionary<string, SystemTypeSemanticSnapshot> semanticMap, string typeClassName, string categoryName, string typeName)
	{
		if (semanticMap == null)
		{
			return null;
		}
		string key = SystemTypeIdentityService.BuildKey(typeClassName, categoryName, typeName);
		SystemTypeSemanticSnapshot snapshot = null;
		if (semanticMap.TryGetValue(key, out snapshot))
		{
			return snapshot;
		}
		return null;
	}
}
