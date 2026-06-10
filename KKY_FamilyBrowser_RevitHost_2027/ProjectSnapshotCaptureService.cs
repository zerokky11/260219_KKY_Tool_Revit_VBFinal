using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Runtime.CompilerServices;
using Autodesk.Revit.DB;
using Microsoft.VisualBasic.CompilerServices;

public sealed class ProjectSnapshotCaptureService
{
	private sealed class LoadableFamilyDeepCaptureResult
	{
		public LoadableFamilyContentSignatureResult SignatureResult { get; set; }

		public List<StandardFamilyParameterSnapshotItem> Parameters { get; set; }

		public bool ParametersCaptured { get; set; }
	}

	[CompilerGenerated]
	internal sealed class _Closure_0024__15_002D0
	{
		public Document _0024VB_0024Local_doc;

		public Func<ElementId, ElementType> _0024I2;

		public _Closure_0024__15_002D0(_Closure_0024__15_002D0 arg0)
		{
			if (arg0 != null)
			{
				_0024VB_0024Local_doc = arg0._0024VB_0024Local_doc;
			}
		}

		[SpecialName]
		internal ElementType _Lambda_0024__2(ElementId id)
		{
			Element element = _0024VB_0024Local_doc.GetElement(id);
			return (ElementType)(object)((element is ElementType) ? element : null);
		}
	}

	[CompilerGenerated]
	internal sealed class _Closure_0024__16_002D0
	{
		public HashSet<string> _0024VB_0024Local_requestedNames;

		public Document _0024VB_0024Local_doc;

		public Func<ElementId, ElementType> _0024I5;

		public _Closure_0024__16_002D0(_Closure_0024__16_002D0 arg0)
		{
			if (arg0 != null)
			{
				_0024VB_0024Local_requestedNames = arg0._0024VB_0024Local_requestedNames;
				_0024VB_0024Local_doc = arg0._0024VB_0024Local_doc;
			}
		}

		[SpecialName]
		internal bool _Lambda_0024__3(Family x)
		{
			return _0024VB_0024Local_requestedNames.Contains(Normalize(((Element)x).Name));
		}

		[SpecialName]
		internal ElementType _Lambda_0024__5(ElementId id)
		{
			Element element = _0024VB_0024Local_doc.GetElement(id);
			return (ElementType)(object)((element is ElementType) ? element : null);
		}
	}

	[CompilerGenerated]
	internal sealed class _Closure_0024__5_002D0
	{
		public Document _0024VB_0024Local_doc;

		public Action<int, int, string> _0024VB_0024Local_progress;

		public Func<ElementId, ElementType> _0024I2;

		public _Closure_0024__5_002D0(_Closure_0024__5_002D0 arg0)
		{
			if (arg0 != null)
			{
				_0024VB_0024Local_doc = arg0._0024VB_0024Local_doc;
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
		internal void _Lambda_0024__6(int current, int total, string message)
		{
			ReportProgress(_0024VB_0024Local_progress, checked(72 + (int)Math.Round((double)current / (double)Math.Max(1, total) * 5.0)), 100, message);
		}
	}

	private static readonly HashSet<string> AllowedSystemTypeNames = new HashSet<string>(new string[21]
	{
		"WallType", "FloorType", "RoofType", "CeilingType", "StairsType", "RailingType", "DuctType", "PipeType", "FlexDuctType", "FlexPipeType",
		"DuctSystemType", "PipingSystemType", "MechanicalSystemType", "ElectricalSystemType", "CableTrayType", "ConduitType", "WireType", "DuctInsulationType", "PipeInsulationType", "DuctLiningType",
		"MullionType"
	}, StringComparer.OrdinalIgnoreCase);

	private static readonly HashSet<string> RoutingAwareSystemTypeNames = new HashSet<string>(new string[6] { "DuctType", "PipeType", "FlexDuctType", "FlexPipeType", "CableTrayType", "ConduitType" }, StringComparer.OrdinalIgnoreCase);

	private ProjectSnapshotCaptureService()
	{
	}

	public static ProjectContentSnapshot Capture(Document doc, IDictionary<string, string> loadableContentFingerprintCache = null, bool includeDeepLoadableContent = false, Action<int, int, string> progress = null, string fingerprintDebugRunFolder = "", FamilyThumbnailConstraintDialogGuard dialogGuard = null)
	{
		//IL_00e4: Unknown result type (might be due to invalid IL or missing references)
		//IL_06c7: Unknown result type (might be due to invalid IL or missing references)
		_Closure_0024__5_002D0 arg = default(_Closure_0024__5_002D0);
		_Closure_0024__5_002D0 CS_0024_003C_003E8__locals25 = new _Closure_0024__5_002D0(arg);
		CS_0024_003C_003E8__locals25._0024VB_0024Local_doc = doc;
		CS_0024_003C_003E8__locals25._0024VB_0024Local_progress = progress;
		ProjectContentSnapshot snapshot = new ProjectContentSnapshot
		{
			DocumentTitle = (CS_0024_003C_003E8__locals25._0024VB_0024Local_doc.Title ?? string.Empty),
			DocumentPath = ProjectSnapshotStore.ResolveProjectIdentityPath(CS_0024_003C_003E8__locals25._0024VB_0024Local_doc),
			CapturedAtUtc = DateTime.UtcNow.ToString("O", CultureInfo.InvariantCulture),
			RevitVersion = CS_0024_003C_003E8__locals25._0024VB_0024Local_doc.Application.VersionNumber
		};
		IDictionary<string, string> effectiveLoadableContentFingerprintCache = loadableContentFingerprintCache ?? new Dictionary<string, string>(StringComparer.Ordinal);
		Dictionary<string, SystemTypeSemanticSnapshot> semanticMap = new Dictionary<string, SystemTypeSemanticSnapshot>(StringComparer.Ordinal);
		ReportProgress(CS_0024_003C_003E8__locals25._0024VB_0024Local_progress, 18, 100, T("Counting current model family instances...", "현재 모델 패밀리 인스턴스 수 계산 중..."));
		Dictionary<int, int> familyInstanceCounts = BuildFamilyInstanceCountMap(CS_0024_003C_003E8__locals25._0024VB_0024Local_doc);
		ReportProgress(CS_0024_003C_003E8__locals25._0024VB_0024Local_progress, 25, 100, T("Collecting current model loadable families...", "현재 모델 로더블 패밀리 수집 중..."));
		List<Family> loadableFamilies = (from Family x in (IEnumerable)new FilteredElementCollector(CS_0024_003C_003E8__locals25._0024VB_0024Local_doc).OfClass(typeof(Family))
			where FamilyBrowserFamilyClassificationService.IsBrowserLoadableFamily(x)
			select x).OrderBy<Family, string>([SpecialName] (Family x) => Normalize(((Element)x).Name), StringComparer.Ordinal).ToList();
		int signatureWrittenCount = 0;
		int signatureMissingCount = 0;
		int fingerprintWrittenCount = 0;
		int fingerprintMissingCount = 0;
		int signatureFileMissingCount = 0;
		int loadableTotal = Math.Max(1, loadableFamilies.Count);
		checked
		{
			int num = loadableFamilies.Count - 1;
			for (int loadableIndex = 0; loadableIndex <= num; loadableIndex++)
			{
				Family family = loadableFamilies[loadableIndex];
				ReportProgress(CS_0024_003C_003E8__locals25._0024VB_0024Local_progress, 28 + (int)Math.Round((double)loadableIndex / (double)loadableTotal * 42.0), 100, T("Reading current model family ", "현재 모델 패밀리 읽는 중 ") + (loadableIndex + 1).ToString(CultureInfo.InvariantCulture) + "/" + loadableFamilies.Count.ToString(CultureInfo.InvariantCulture) + ": " + (((Element)family).Name ?? string.Empty));
				List<string> typeNames = (from x in family.GetFamilySymbolIds().Select([SpecialName] (ElementId id) =>
					{
						Element element = CS_0024_003C_003E8__locals25._0024VB_0024Local_doc.GetElement(id);
						return (ElementType)(object)((element is ElementType) ? element : null);
					})
					where x != null
					select ((Element)x).Name ?? string.Empty).OrderBy<string, string>([SpecialName] (string x) => Normalize(x), StringComparer.Ordinal).ToList();
				string categoryName = FamilyBrowserFamilyClassificationService.ResolveCategoryName(family);
				string categoryId = FamilyBrowserFamilyClassificationService.ResolveCategoryId(family);
				string categoryGroup = FamilyBrowserFamilyClassificationService.ResolveCategoryGroup(family);
				string familyName = ((Element)family).Name ?? string.Empty;
				string contentFingerprint = string.Empty;
				string contentSignatureDebugPath = string.Empty;
				string contentFingerprintFailureReason = string.Empty;
				string contentSignatureFileFailureReason = string.Empty;
				List<StandardFamilyParameterSnapshotItem> capturedParameters = null;
				bool capturedParametersFromFamilyDocument = false;
				int dialogRecordStart = dialogGuard?.RecordCount ?? 0;
				dialogGuard?.SetCurrentFamily(categoryName, familyName);
				try
				{
					if (!string.IsNullOrWhiteSpace(fingerprintDebugRunFolder))
					{
						LoadableFamilyDeepCaptureResult deepCapture = (includeDeepLoadableContent ? CaptureLoadableFamilyDeepData(CS_0024_003C_003E8__locals25._0024VB_0024Local_doc, family) : null);
						LoadableFamilyContentSignatureResult signatureResult = ((deepCapture != null) ? deepCapture.SignatureResult : LoadableFamilyContentSignatureService.BuildResult(CS_0024_003C_003E8__locals25._0024VB_0024Local_doc, family, includeDeepLoadableContent));
						if (deepCapture != null && deepCapture.ParametersCaptured)
						{
							capturedParameters = deepCapture.Parameters;
							capturedParametersFromFamilyDocument = true;
						}
						contentFingerprint = signatureResult.Fingerprint ?? string.Empty;
						if (string.IsNullOrWhiteSpace(contentFingerprint))
						{
							contentFingerprintFailureReason = ResolveFingerprintFailureReason(signatureResult, "Project family fingerprint was empty.");
						}
						string dialogFailureReason = ResolveDialogCancelFingerprintFailureReason(dialogGuard, dialogRecordStart);
						if (!string.IsNullOrWhiteSpace(dialogFailureReason))
						{
							contentFingerprintFailureReason = dialogFailureReason;
						}
						if (string.IsNullOrWhiteSpace(contentFingerprint) && !string.IsNullOrWhiteSpace(contentFingerprintFailureReason))
						{
							signatureResult = LoadableFamilyContentSignatureService.BuildDiagnosticFailureResult(includeDeepLoadableContent ? "Precise" : "Fast", contentFingerprintFailureReason, family);
						}
						contentSignatureDebugPath = FingerprintDebugSignatureStore.SaveLoadableSignature(fingerprintDebugRunFolder, "project", snapshot.DocumentTitle, categoryName, familyName, contentFingerprint, signatureResult, ref contentSignatureFileFailureReason);
						if (string.IsNullOrWhiteSpace(contentSignatureDebugPath) && !string.IsNullOrWhiteSpace(contentSignatureFileFailureReason))
						{
							contentFingerprintFailureReason = AppendFailureReason(contentFingerprintFailureReason, contentSignatureFileFailureReason);
						}
						string cacheKey = LoadableFamilyContentSignatureService.BuildCacheKey(CS_0024_003C_003E8__locals25._0024VB_0024Local_doc, family, includeDeepLoadableContent);
						if (!string.IsNullOrWhiteSpace(cacheKey))
						{
							effectiveLoadableContentFingerprintCache[cacheKey] = contentFingerprint;
						}
						if (!string.IsNullOrWhiteSpace(contentSignatureDebugPath))
						{
							signatureWrittenCount++;
						}
						else
						{
							signatureFileMissingCount++;
						}
						if (includeDeepLoadableContent)
						{
							if (string.IsNullOrWhiteSpace(contentFingerprint))
							{
								fingerprintMissingCount++;
								signatureMissingCount++;
							}
							else
							{
								fingerprintWrittenCount++;
							}
						}
					}
					else
					{
						LoadableFamilyDeepCaptureResult deepCapture2 = (includeDeepLoadableContent ? CaptureLoadableFamilyDeepData(CS_0024_003C_003E8__locals25._0024VB_0024Local_doc, family) : null);
						LoadableFamilyContentSignatureResult signatureResult2 = ((deepCapture2 != null) ? deepCapture2.SignatureResult : LoadableFamilyContentSignatureService.BuildResult(CS_0024_003C_003E8__locals25._0024VB_0024Local_doc, family, includeDeepLoadableContent));
						if (deepCapture2 != null && deepCapture2.ParametersCaptured)
						{
							capturedParameters = deepCapture2.Parameters;
							capturedParametersFromFamilyDocument = true;
						}
						contentFingerprint = signatureResult2.Fingerprint ?? string.Empty;
						if (string.IsNullOrWhiteSpace(contentFingerprint))
						{
							contentFingerprintFailureReason = ResolveFingerprintFailureReason(signatureResult2, "Project family fingerprint was empty.");
						}
						string cacheKey2 = LoadableFamilyContentSignatureService.BuildCacheKey(CS_0024_003C_003E8__locals25._0024VB_0024Local_doc, family, includeDeepLoadableContent);
						if (!string.IsNullOrWhiteSpace(cacheKey2))
						{
							effectiveLoadableContentFingerprintCache[cacheKey2] = contentFingerprint;
						}
						string dialogFailureReason2 = ResolveDialogCancelFingerprintFailureReason(dialogGuard, dialogRecordStart);
						if (!string.IsNullOrWhiteSpace(dialogFailureReason2))
						{
							contentFingerprintFailureReason = dialogFailureReason2;
						}
						if (includeDeepLoadableContent && string.IsNullOrWhiteSpace(contentFingerprint))
						{
							fingerprintMissingCount++;
							signatureMissingCount++;
						}
						else if (includeDeepLoadableContent)
						{
							fingerprintWrittenCount++;
						}
					}
				}
				finally
				{
					dialogGuard?.ClearCurrentFamily();
				}
				snapshot.LoadableFamilies.Add(new ProjectLoadableFamilySnapshotItem
				{
					FamilyName = (((Element)family).Name ?? string.Empty),
					CategoryName = categoryName,
					CategoryId = categoryId,
					CategoryGroup = categoryGroup,
					TypeCount = typeNames.Count,
					InstanceCount = ResolveFamilyInstanceCount(familyInstanceCounts, family),
					TypeNames = typeNames,
					Parameters = ResolveLoadableFamilyParameterSnapshots(CS_0024_003C_003E8__locals25._0024VB_0024Local_doc, family, capturedParameters, capturedParametersFromFamilyDocument),
					ContentFingerprint = contentFingerprint,
					ContentSignatureDebugPath = contentSignatureDebugPath,
					ContentFingerprintFailureReason = contentFingerprintFailureReason,
					UniqueId = (((Element)family).UniqueId ?? string.Empty),
					IsShared = ResolveIsShared(family)
				});
			}
			ReportProgress(CS_0024_003C_003E8__locals25._0024VB_0024Local_progress, 72, 100, T("Building current model system type dependency map...", "현재 모델 시스템 타입 의존성 맵 작성 중..."));
			semanticMap = SystemTypeSemanticFingerprintCatalogService.BuildSnapshotMap(CS_0024_003C_003E8__locals25._0024VB_0024Local_doc, "project|" + snapshot.DocumentTitle, effectiveLoadableContentFingerprintCache, includeDeepLoadableContent, [SpecialName] (int current, int total, string message) =>
			{
				ReportProgress(CS_0024_003C_003E8__locals25._0024VB_0024Local_progress, 72 + (int)Math.Round((double)current / (double)Math.Max(1, total) * 5.0), 100, message);
			});
			ReportProgress(CS_0024_003C_003E8__locals25._0024VB_0024Local_progress, 78, 100, T("Collecting current model system types...", "현재 모델 시스템 타입 수집 중..."));
			List<ElementType> systemTypes = (from ElementType x in (IEnumerable)new FilteredElementCollector(CS_0024_003C_003E8__locals25._0024VB_0024Local_doc).WhereElementIsElementType()
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
				ReportProgress(CS_0024_003C_003E8__locals25._0024VB_0024Local_progress, 78 + (int)Math.Round((double)systemIndex / (double)systemTotal * 14.0), 100, T("Reading current model system type ", "현재 모델 시스템 타입 읽는 중 ") + (systemIndex + 1).ToString(CultureInfo.InvariantCulture) + "/" + systemTypes.Count.ToString(CultureInfo.InvariantCulture) + ": " + ((object)systemType).GetType().Name + " / " + (((Element)systemType).Name ?? string.Empty));
				snapshot.SystemTypes.Add(new ProjectSystemTypeSnapshotItem
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
					CompoundStructureSignature = ((semanticSnapshot == null) ? string.Empty : semanticSnapshot.CompoundStructureSignature)
				});
			}
			ReportProgress(CS_0024_003C_003E8__locals25._0024VB_0024Local_progress, 96, 100, T("Summarizing current model scan...", "현재 모델 스캔 요약 중..."));
			snapshot.LoadableSignatureFailures = BuildLoadableSignatureFailures(snapshot.LoadableFamilies);
			if (includeDeepLoadableContent && fingerprintWrittenCount == 0 && fingerprintMissingCount == 0)
			{
				fingerprintWrittenCount = snapshot.LoadableFamilies.Where([SpecialName] (ProjectLoadableFamilySnapshotItem x) => x != null && !string.IsNullOrWhiteSpace(x.ContentFingerprint)).Count();
				fingerprintMissingCount = snapshot.LoadableSignatureFailures.Where([SpecialName] (ProjectLoadableSignatureFailureItem x) => x != null && (x.FailureKind ?? string.Empty).IndexOf("FingerprintMissing", StringComparison.OrdinalIgnoreCase) >= 0).Count();
				signatureMissingCount = fingerprintMissingCount;
			}
			snapshot.Summary = new ProjectContentSnapshotSummary
			{
				LoadableFamilyCount = snapshot.LoadableFamilies.Count,
				LoadableTypeCount = snapshot.LoadableFamilies.Sum([SpecialName] (ProjectLoadableFamilySnapshotItem x) => x.TypeCount),
				LoadableFingerprintWrittenCount = fingerprintWrittenCount,
				LoadableFingerprintMissingCount = fingerprintMissingCount,
				LoadableSignatureWrittenCount = signatureWrittenCount,
				LoadableSignatureMissingCount = signatureMissingCount,
				LoadableSignatureFileMissingCount = signatureFileMissingCount,
				SystemTypeCount = snapshot.SystemTypes.Count,
				SystemTypeClassCount = snapshot.SystemTypes.Select([SpecialName] (ProjectSystemTypeSnapshotItem x) => x.TypeClassName).Distinct<string>(StringComparer.OrdinalIgnoreCase).Count()
			};
			return snapshot;
		}
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

	private static LoadableFamilyDeepCaptureResult CaptureLoadableFamilyDeepData(Document doc, Family family)
	{
		LoadableFamilyDeepCaptureResult result = new LoadableFamilyDeepCaptureResult
		{
			Parameters = new List<StandardFamilyParameterSnapshotItem>()
		};
		LoadableFamilyDeepCaptureResult CaptureLoadableFamilyDeepData;
		if (doc == null || family == null)
		{
			result.SignatureResult = LoadableFamilyContentSignatureService.BuildDiagnosticFailureResult("Precise", "Host document or family was empty.", family);
			CaptureLoadableFamilyDeepData = result;
		}
		else if (!CanEditFamilyForDeepCapture(family))
		{
			result.SignatureResult = LoadableFamilyContentSignatureService.BuildDiagnosticFailureResult("Precise", "Family is not editable or is in-place.", family);
			CaptureLoadableFamilyDeepData = result;
		}
		else
		{
			Document familyDoc = null;
			try
			{
				familyDoc = doc.EditFamily(family);
				if (familyDoc == null)
				{
					result.SignatureResult = LoadableFamilyContentSignatureService.BuildDiagnosticFailureResult("Precise", "EditFamily returned no document.", family);
					CaptureLoadableFamilyDeepData = result;
				}
				else
				{
					result.SignatureResult = LoadableFamilyContentSignatureService.BuildResultFromOpenFamilyDocument(family, familyDoc);
					result.Parameters = FamilyDocumentParameterCaptureService.Capture(familyDoc);
					result.ParametersCaptured = true;
					CaptureLoadableFamilyDeepData = result;
				}
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				result.SignatureResult = LoadableFamilyContentSignatureService.BuildDiagnosticFailureResult("Precise", "EditFamily project precise capture failed: " + ex2.GetType().Name + " - " + ex2.Message, family);
				CaptureLoadableFamilyDeepData = result;
				ProjectData.ClearProjectError();
			}
			finally
			{
				if (familyDoc != null)
				{
					try
					{
						familyDoc.Close(false);
					}
					catch (Exception projectError)
					{
						ProjectData.SetProjectError(projectError);
						ProjectData.ClearProjectError();
					}
				}
			}
		}
		return CaptureLoadableFamilyDeepData;
	}

	private static bool CanEditFamilyForDeepCapture(Family family)
	{
		bool CanEditFamilyForDeepCapture;
		try
		{
			CanEditFamilyForDeepCapture = family != null && !family.IsInPlace && family.IsEditable;
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			CanEditFamilyForDeepCapture = false;
			ProjectData.ClearProjectError();
		}
		return CanEditFamilyForDeepCapture;
	}

	private static List<StandardFamilyParameterSnapshotItem> ResolveLoadableFamilyParameterSnapshots(Document doc, Family family, List<StandardFamilyParameterSnapshotItem> familyDocumentParameters, bool familyDocumentParametersCaptured)
	{
		if (familyDocumentParametersCaptured)
		{
			return familyDocumentParameters ?? new List<StandardFamilyParameterSnapshotItem>();
		}
		return CaptureLoadableFamilyParameters(doc, family);
	}

	private static string T(string englishText, string koreanText)
	{
		return FamilyBrowserLanguageService.Text(englishText, koreanText);
	}

	private static string ResolveDialogCancelFingerprintFailureReason(FamilyThumbnailConstraintDialogGuard dialogGuard, int dialogRecordStart)
	{
		if (dialogGuard == null)
		{
			return string.Empty;
		}
		return FamilyThumbnailConstraintDialogGuard.BuildFingerprintCanceledReason(dialogGuard.GetRecordsSince(dialogRecordStart));
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

	public static List<ProjectLoadableSignatureFailureItem> BuildLoadableSignatureFailures(IEnumerable<ProjectLoadableFamilySnapshotItem> items)
	{
		List<ProjectLoadableSignatureFailureItem> failures = new List<ProjectLoadableSignatureFailureItem>();
		if (items == null)
		{
			return failures;
		}
		foreach (ProjectLoadableFamilySnapshotItem item in items)
		{
			if (item != null)
			{
				List<string> kinds = new List<string>();
				if (string.IsNullOrWhiteSpace(item.ContentFingerprint))
				{
					kinds.Add("FingerprintMissing");
				}
				if (!string.IsNullOrWhiteSpace(item.ContentFingerprintFailureReason) && item.ContentFingerprintFailureReason.IndexOf("Signature diagnostic file", StringComparison.OrdinalIgnoreCase) >= 0 && string.IsNullOrWhiteSpace(item.ContentSignatureDebugPath))
				{
					kinds.Add("SignatureFileMissing");
				}
				if (kinds.Count != 0)
				{
					failures.Add(new ProjectLoadableSignatureFailureItem
					{
						FamilyName = (item.FamilyName ?? string.Empty),
						CategoryName = (item.CategoryName ?? string.Empty),
						CategoryId = (item.CategoryId ?? string.Empty),
						CategoryGroup = (item.CategoryGroup ?? string.Empty),
						TypeCount = item.TypeCount,
						InstanceCount = item.InstanceCount,
						FailureKind = string.Join(";", kinds.Distinct<string>(StringComparer.OrdinalIgnoreCase)),
						Reason = (string.IsNullOrWhiteSpace(item.ContentFingerprintFailureReason) ? "Fingerprint was empty without a recorded reason." : item.ContentFingerprintFailureReason),
						ContentFingerprint = (item.ContentFingerprint ?? string.Empty),
						ContentSignatureDebugPath = (item.ContentSignatureDebugPath ?? string.Empty),
						UniqueId = (item.UniqueId ?? string.Empty),
						IsShared = item.IsShared
					});
				}
			}
		}
		return failures;
	}

	private static string AppendFailureReason(string currentReason, string additionalReason)
	{
		if (string.IsNullOrWhiteSpace(additionalReason))
		{
			return currentReason ?? string.Empty;
		}
		if (string.IsNullOrWhiteSpace(currentReason))
		{
			return additionalReason.Trim();
		}
		if (currentReason.IndexOf(additionalReason.Trim(), StringComparison.OrdinalIgnoreCase) >= 0)
		{
			return currentReason;
		}
		return currentReason.Trim() + " " + additionalReason.Trim();
	}

	public static ProjectContentSnapshot CaptureLoadableFamilies(Document doc, IDictionary<string, string> loadableContentFingerprintCache = null, bool includeDeepLoadableContent = false)
	{
		//IL_0095: Unknown result type (might be due to invalid IL or missing references)
		_Closure_0024__15_002D0 arg = default(_Closure_0024__15_002D0);
		_Closure_0024__15_002D0 CS_0024_003C_003E8__locals11 = new _Closure_0024__15_002D0(arg);
		CS_0024_003C_003E8__locals11._0024VB_0024Local_doc = doc;
		ProjectContentSnapshot snapshot = new ProjectContentSnapshot
		{
			DocumentTitle = (CS_0024_003C_003E8__locals11._0024VB_0024Local_doc.Title ?? string.Empty),
			DocumentPath = ProjectSnapshotStore.ResolveProjectIdentityPath(CS_0024_003C_003E8__locals11._0024VB_0024Local_doc),
			CapturedAtUtc = DateTime.UtcNow.ToString("O", CultureInfo.InvariantCulture),
			RevitVersion = CS_0024_003C_003E8__locals11._0024VB_0024Local_doc.Application.VersionNumber
		};
		IDictionary<string, string> effectiveLoadableContentFingerprintCache = loadableContentFingerprintCache ?? new Dictionary<string, string>(StringComparer.Ordinal);
		Dictionary<int, int> familyInstanceCounts = BuildFamilyInstanceCountMap(CS_0024_003C_003E8__locals11._0024VB_0024Local_doc);
		IOrderedEnumerable<Family> loadableFamilies = (from Family x in (IEnumerable)new FilteredElementCollector(CS_0024_003C_003E8__locals11._0024VB_0024Local_doc).OfClass(typeof(Family))
			where FamilyBrowserFamilyClassificationService.IsBrowserLoadableFamily(x)
			select x).OrderBy<Family, string>([SpecialName] (Family x) => Normalize(((Element)x).Name), StringComparer.Ordinal);
		foreach (Family family in loadableFamilies)
		{
			List<string> typeNames = (from x in family.GetFamilySymbolIds().Select([SpecialName] (ElementId id) =>
				{
					Element element = CS_0024_003C_003E8__locals11._0024VB_0024Local_doc.GetElement(id);
					return (ElementType)(object)((element is ElementType) ? element : null);
				})
				where x != null
				select ((Element)x).Name ?? string.Empty).OrderBy<string, string>([SpecialName] (string x) => Normalize(x), StringComparer.Ordinal).ToList();
			LoadableFamilyDeepCaptureResult deepCapture = (includeDeepLoadableContent ? CaptureLoadableFamilyDeepData(CS_0024_003C_003E8__locals11._0024VB_0024Local_doc, family) : null);
			List<StandardFamilyParameterSnapshotItem> capturedParameters = null;
			bool capturedParametersFromFamilyDocument = false;
			string contentFingerprint;
			if (deepCapture != null)
			{
				contentFingerprint = deepCapture.SignatureResult?.Fingerprint ?? string.Empty;
				capturedParameters = deepCapture.Parameters;
				capturedParametersFromFamilyDocument = deepCapture.ParametersCaptured;
				string cacheKey = LoadableFamilyContentSignatureService.BuildCacheKey(CS_0024_003C_003E8__locals11._0024VB_0024Local_doc, family, includeDeepLoadableContent);
				if (!string.IsNullOrWhiteSpace(cacheKey))
				{
					effectiveLoadableContentFingerprintCache[cacheKey] = contentFingerprint;
				}
			}
			else
			{
				contentFingerprint = LoadableFamilyContentSignatureService.Build(CS_0024_003C_003E8__locals11._0024VB_0024Local_doc, family, effectiveLoadableContentFingerprintCache, includeDeepLoadableContent);
			}
			snapshot.LoadableFamilies.Add(new ProjectLoadableFamilySnapshotItem
			{
				FamilyName = (((Element)family).Name ?? string.Empty),
				CategoryName = FamilyBrowserFamilyClassificationService.ResolveCategoryName(family),
				CategoryId = FamilyBrowserFamilyClassificationService.ResolveCategoryId(family),
				CategoryGroup = FamilyBrowserFamilyClassificationService.ResolveCategoryGroup(family),
				TypeCount = typeNames.Count,
				InstanceCount = ResolveFamilyInstanceCount(familyInstanceCounts, family),
				TypeNames = typeNames,
				Parameters = ResolveLoadableFamilyParameterSnapshots(CS_0024_003C_003E8__locals11._0024VB_0024Local_doc, family, capturedParameters, capturedParametersFromFamilyDocument),
				ContentFingerprint = contentFingerprint,
				UniqueId = (((Element)family).UniqueId ?? string.Empty),
				IsShared = ResolveIsShared(family)
			});
		}
		snapshot.Summary = new ProjectContentSnapshotSummary
		{
			LoadableFamilyCount = snapshot.LoadableFamilies.Count,
			LoadableTypeCount = snapshot.LoadableFamilies.Sum([SpecialName] (ProjectLoadableFamilySnapshotItem x) => x.TypeCount),
			SystemTypeCount = 0,
			SystemTypeClassCount = 0
		};
		return snapshot;
	}

	public static ProjectContentSnapshot CaptureLoadableFamiliesByNames(Document doc, IEnumerable<string> familyNames, IDictionary<string, string> loadableContentFingerprintCache = null, bool includeDeepLoadableContent = false, bool includeInstanceCounts = true, Action<int, int, string> progress = null)
	{
		//IL_012c: Unknown result type (might be due to invalid IL or missing references)
		_Closure_0024__16_002D0 arg = default(_Closure_0024__16_002D0);
		_Closure_0024__16_002D0 CS_0024_003C_003E8__locals14 = new _Closure_0024__16_002D0(arg);
		CS_0024_003C_003E8__locals14._0024VB_0024Local_doc = doc;
		ProjectContentSnapshot snapshot = new ProjectContentSnapshot
		{
			DocumentTitle = (CS_0024_003C_003E8__locals14._0024VB_0024Local_doc.Title ?? string.Empty),
			DocumentPath = ProjectSnapshotStore.ResolveProjectIdentityPath(CS_0024_003C_003E8__locals14._0024VB_0024Local_doc),
			CapturedAtUtc = DateTime.UtcNow.ToString("O", CultureInfo.InvariantCulture),
			RevitVersion = CS_0024_003C_003E8__locals14._0024VB_0024Local_doc.Application.VersionNumber
		};
		CS_0024_003C_003E8__locals14._0024VB_0024Local_requestedNames = new HashSet<string>(from x in familyNames ?? Enumerable.Empty<string>()
			where !string.IsNullOrWhiteSpace(x)
			select Normalize(x), StringComparer.Ordinal);
		if (CS_0024_003C_003E8__locals14._0024VB_0024Local_requestedNames.Count == 0)
		{
			snapshot.Summary = new ProjectContentSnapshotSummary();
			return snapshot;
		}
		IDictionary<string, string> effectiveLoadableContentFingerprintCache = loadableContentFingerprintCache ?? new Dictionary<string, string>(StringComparer.Ordinal);
		Dictionary<int, int> familyInstanceCounts = (includeInstanceCounts ? BuildFamilyInstanceCountMap(CS_0024_003C_003E8__locals14._0024VB_0024Local_doc) : new Dictionary<int, int>());
		List<Family> loadableFamilies = (from Family x in (IEnumerable)new FilteredElementCollector(CS_0024_003C_003E8__locals14._0024VB_0024Local_doc).OfClass(typeof(Family))
			where FamilyBrowserFamilyClassificationService.IsBrowserLoadableFamily(x)
			where CS_0024_003C_003E8__locals14._0024VB_0024Local_requestedNames.Contains(Normalize(((Element)x).Name))
			select x).OrderBy<Family, string>([SpecialName] (Family x) => Normalize(((Element)x).Name), StringComparer.Ordinal).ToList();
		int total = Math.Max(1, loadableFamilies.Count);
		checked
		{
			int num = loadableFamilies.Count - 1;
			for (int index = 0; index <= num; index++)
			{
				Family family = loadableFamilies[index];
				ReportProgress(progress, index + 1, total, "Reading dependency family " + (index + 1).ToString(CultureInfo.InvariantCulture) + "/" + total.ToString(CultureInfo.InvariantCulture) + ": " + (((Element)family).Name ?? string.Empty));
				List<string> typeNames = (from x in family.GetFamilySymbolIds().Select([SpecialName] (ElementId id) =>
					{
						Element element = CS_0024_003C_003E8__locals14._0024VB_0024Local_doc.GetElement(id);
						return (ElementType)(object)((element is ElementType) ? element : null);
					})
					where x != null
					select ((Element)x).Name ?? string.Empty).OrderBy<string, string>([SpecialName] (string x) => Normalize(x), StringComparer.Ordinal).ToList();
				LoadableFamilyDeepCaptureResult deepCapture = (includeDeepLoadableContent ? CaptureLoadableFamilyDeepData(CS_0024_003C_003E8__locals14._0024VB_0024Local_doc, family) : null);
				List<StandardFamilyParameterSnapshotItem> capturedParameters = null;
				bool capturedParametersFromFamilyDocument = false;
				string contentFingerprint;
				if (deepCapture != null)
				{
					contentFingerprint = deepCapture.SignatureResult?.Fingerprint ?? string.Empty;
					capturedParameters = deepCapture.Parameters;
					capturedParametersFromFamilyDocument = deepCapture.ParametersCaptured;
					string cacheKey = LoadableFamilyContentSignatureService.BuildCacheKey(CS_0024_003C_003E8__locals14._0024VB_0024Local_doc, family, includeDeepLoadableContent);
					if (!string.IsNullOrWhiteSpace(cacheKey))
					{
						effectiveLoadableContentFingerprintCache[cacheKey] = contentFingerprint;
					}
				}
				else
				{
					contentFingerprint = LoadableFamilyContentSignatureService.Build(CS_0024_003C_003E8__locals14._0024VB_0024Local_doc, family, effectiveLoadableContentFingerprintCache, includeDeepLoadableContent);
				}
				snapshot.LoadableFamilies.Add(new ProjectLoadableFamilySnapshotItem
				{
					FamilyName = (((Element)family).Name ?? string.Empty),
					CategoryName = FamilyBrowserFamilyClassificationService.ResolveCategoryName(family),
					CategoryId = FamilyBrowserFamilyClassificationService.ResolveCategoryId(family),
					CategoryGroup = FamilyBrowserFamilyClassificationService.ResolveCategoryGroup(family),
					TypeCount = typeNames.Count,
					InstanceCount = ResolveFamilyInstanceCount(familyInstanceCounts, family),
					TypeNames = typeNames,
					Parameters = ResolveLoadableFamilyParameterSnapshots(CS_0024_003C_003E8__locals14._0024VB_0024Local_doc, family, capturedParameters, capturedParametersFromFamilyDocument),
					ContentFingerprint = contentFingerprint,
					UniqueId = (((Element)family).UniqueId ?? string.Empty),
					IsShared = ResolveIsShared(family)
				});
			}
			snapshot.Summary = new ProjectContentSnapshotSummary
			{
				LoadableFamilyCount = snapshot.LoadableFamilies.Count,
				LoadableTypeCount = snapshot.LoadableFamilies.Sum([SpecialName] (ProjectLoadableFamilySnapshotItem x) => x.TypeCount),
				SystemTypeCount = 0,
				SystemTypeClassCount = 0
			};
			return snapshot;
		}
	}

	public static ProjectContentSnapshot CaptureSystemTypes(Document doc, IDictionary<string, string> loadableContentFingerprintCache = null, bool includeDeepLoadableContent = false)
	{
		//IL_0081: Unknown result type (might be due to invalid IL or missing references)
		ProjectContentSnapshot snapshot = new ProjectContentSnapshot
		{
			DocumentTitle = (doc.Title ?? string.Empty),
			DocumentPath = ProjectSnapshotStore.ResolveProjectIdentityPath(doc),
			CapturedAtUtc = DateTime.UtcNow.ToString("O", CultureInfo.InvariantCulture),
			RevitVersion = doc.Application.VersionNumber
		};
		IDictionary<string, string> effectiveLoadableContentFingerprintCache = loadableContentFingerprintCache ?? new Dictionary<string, string>(StringComparer.Ordinal);
		Dictionary<string, SystemTypeSemanticSnapshot> semanticMap = SystemTypeSemanticFingerprintCatalogService.BuildSnapshotMap(doc, "project-system|" + snapshot.DocumentTitle, effectiveLoadableContentFingerprintCache, includeDeepLoadableContent);
		IOrderedEnumerable<ElementType> systemTypes = (from ElementType x in (IEnumerable)new FilteredElementCollector(doc).WhereElementIsElementType()
			where x != null
			where !(x is FamilySymbol)
			where AllowedSystemTypeNames.Contains(((object)x).GetType().Name)
			select x).OrderBy<ElementType, string>([SpecialName] (ElementType x) => ((object)x).GetType().Name, StringComparer.Ordinal).ThenBy<ElementType, string>([SpecialName] (ElementType x) => Normalize(((Element)x).Name), StringComparer.Ordinal);
		foreach (ElementType systemType in systemTypes)
		{
			SystemTypeSemanticSnapshot semanticSnapshot = ResolveSemanticSnapshot(semanticMap, ((object)systemType).GetType().Name, ResolveCategoryName((Element)(object)systemType), ((Element)systemType).Name ?? string.Empty);
			snapshot.SystemTypes.Add(new ProjectSystemTypeSnapshotItem
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
				CompoundStructureSignature = ((semanticSnapshot == null) ? string.Empty : semanticSnapshot.CompoundStructureSignature)
			});
		}
		snapshot.Summary = new ProjectContentSnapshotSummary
		{
			LoadableFamilyCount = 0,
			LoadableTypeCount = 0,
			SystemTypeCount = snapshot.SystemTypes.Count,
			SystemTypeClassCount = snapshot.SystemTypes.Select([SpecialName] (ProjectSystemTypeSnapshotItem x) => x.TypeClassName).Distinct<string>(StringComparer.OrdinalIgnoreCase).Count()
		};
		return snapshot;
	}

	private static Dictionary<int, int> BuildFamilyInstanceCountMap(Document doc)
	{
		//IL_0007: Unknown result type (might be due to invalid IL or missing references)
		Dictionary<int, int> result = new Dictionary<int, int>();
		checked
		{
			try
			{
				foreach (FamilyInstance instance in ((IEnumerable)new FilteredElementCollector(doc).OfClass(typeof(FamilyInstance)).WhereElementIsNotElementType()).Cast<FamilyInstance>())
				{
					if (instance == null || instance.Symbol == null || instance.Symbol.Family == null)
					{
						continue;
					}
					int key = ElementIdKey(((Element)instance.Symbol.Family).Id);
					if (key != int.MinValue)
					{
						if (result.ContainsKey(key))
						{
							result[key]++;
						}
						else
						{
							result[key] = 1;
						}
					}
				}
			}
			catch (Exception projectError)
			{
				ProjectData.SetProjectError(projectError);
				ProjectData.ClearProjectError();
			}
			return result;
		}
	}

	private static int ResolveFamilyInstanceCount(Dictionary<int, int> instanceCounts, Family family)
	{
		if (instanceCounts == null || family == null)
		{
			return 0;
		}
		int key = ElementIdKey(((Element)family).Id);
		if (key == int.MinValue || !instanceCounts.ContainsKey(key))
		{
			return 0;
		}
		return instanceCounts[key];
	}

	private static int ElementIdKey(ElementId id)
	{
		if (id == null)
		{
			return int.MinValue;
		}
		return RevitElementIdCompat.CompatIntegerValue(id);
	}

	public static ProjectContentSnapshot CaptureSystemTypesByKey(Document doc, string systemFamilyKind, string categoryName, string typeName, IDictionary<string, string> loadableContentFingerprintCache = null, bool includeDeepLoadableContent = false)
	{
		ProjectContentSnapshot snapshot = new ProjectContentSnapshot
		{
			DocumentTitle = (doc.Title ?? string.Empty),
			DocumentPath = ProjectSnapshotStore.ResolveProjectIdentityPath(doc),
			CapturedAtUtc = DateTime.UtcNow.ToString("O", CultureInfo.InvariantCulture),
			RevitVersion = doc.Application.VersionNumber
		};
		SystemTypeCatalogSnapshot semanticCatalog = SystemTypeSemanticCaptureService.CaptureSelected(doc, "project-system-selected|" + snapshot.DocumentTitle, systemFamilyKind, categoryName, typeName, loadableContentFingerprintCache ?? new Dictionary<string, string>(StringComparer.Ordinal), includeDeepLoadableContent);
		foreach (SystemTypeSemanticSnapshot semanticType in semanticCatalog.Types)
		{
			ElementType elementType = FindSystemType(doc, semanticType.SystemFamilyKind, semanticType.CategoryName, semanticType.TypeName);
			snapshot.SystemTypes.Add(new ProjectSystemTypeSnapshotItem
			{
				TypeName = semanticType.TypeName,
				CategoryName = semanticType.CategoryName,
				CategoryId = ResolveCategoryId((Element)(object)elementType),
				TypeClassName = semanticType.SystemFamilyKind,
				UniqueId = (((elementType != null) ? ((Element)elementType).UniqueId : null) ?? string.Empty),
				SupportsRoutingDependencies = RoutingAwareSystemTypeNames.Contains(semanticType.SystemFamilyKind),
				SemanticFingerprint = SystemTypeFingerprintService.Compute(semanticType),
				ClassificationCode = (semanticType.ClassificationCode ?? string.Empty),
				SegmentName = (semanticType.SegmentName ?? string.Empty),
				MaterialName = (semanticType.MaterialName ?? string.Empty),
				Shape = (semanticType.Shape ?? string.Empty),
				RoutingPreferenceSignature = (semanticType.RoutingPreferenceSignature ?? string.Empty),
				CompoundStructureSignature = (semanticType.CompoundStructureSignature ?? string.Empty)
			});
		}
		snapshot.Summary = new ProjectContentSnapshotSummary
		{
			LoadableFamilyCount = 0,
			LoadableTypeCount = 0,
			SystemTypeCount = snapshot.SystemTypes.Count,
			SystemTypeClassCount = snapshot.SystemTypes.Select([SpecialName] (ProjectSystemTypeSnapshotItem x) => x.TypeClassName).Distinct<string>(StringComparer.OrdinalIgnoreCase).Count()
		};
		return snapshot;
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

	private static ElementType FindSystemType(Document doc, string systemFamilyKind, string categoryName, string typeName)
	{
		//IL_0040: Unknown result type (might be due to invalid IL or missing references)
		if (doc == null || string.IsNullOrWhiteSpace(systemFamilyKind) || string.IsNullOrWhiteSpace(typeName))
		{
			return null;
		}
		List<ElementType> candidates = (from ElementType x in (IEnumerable)new FilteredElementCollector(doc).WhereElementIsElementType()
			where x != null
			where !(x is FamilySymbol)
			where string.Equals(Normalize(((object)x).GetType().Name), Normalize(systemFamilyKind), StringComparison.Ordinal)
			where string.Equals(Normalize(ResolveElementName((Element)(object)x)), Normalize(typeName), StringComparison.Ordinal)
			select x).ToList();
		return candidates.FirstOrDefault([SpecialName] (ElementType x) => CategoryNamesMatch(ResolveCategoryName((Element)(object)x), categoryName)) ?? candidates.FirstOrDefault();
	}

	private static bool CategoryNamesMatch(string left, string right)
	{
		if (string.IsNullOrWhiteSpace(left) || string.IsNullOrWhiteSpace(right))
		{
			return string.IsNullOrWhiteSpace(left) && string.IsNullOrWhiteSpace(right);
		}
		return string.Equals(Normalize(left), Normalize(right), StringComparison.Ordinal);
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
