using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Security.Cryptography;
using System.Text;
using Autodesk.Revit.DB;
using Microsoft.VisualBasic.CompilerServices;

public sealed class LoadableFamilyContentSignatureService
{
	[CompilerGenerated]
	internal sealed class _Closure_0024__14_002D0
	{
		public bool _0024VB_0024Local_includeAnnotationGraphics;

		public _Closure_0024__14_002D0(_Closure_0024__14_002D0 arg0)
		{
			if (arg0 != null)
			{
				_0024VB_0024Local_includeAnnotationGraphics = arg0._0024VB_0024Local_includeAnnotationGraphics;
			}
		}

		[SpecialName]
		internal bool _Lambda_0024__0(Element x)
		{
			return ShouldCaptureElement(x, _0024VB_0024Local_includeAnnotationGraphics);
		}
	}

	[CompilerGenerated]
	internal sealed class _Closure_0024__24_002D0
	{
		public Parameter _0024VB_0024Local_parameter;

		public _Closure_0024__24_002D0(_Closure_0024__24_002D0 arg0)
		{
			if (arg0 != null)
			{
				_0024VB_0024Local_parameter = arg0._0024VB_0024Local_parameter;
			}
		}

		[SpecialName]
		internal bool _Lambda_0024__2()
		{
			return _0024VB_0024Local_parameter.IsShared;
		}
	}

	private LoadableFamilyContentSignatureService()
	{
	}

	public static string Build(Document hostDocument, Family family, bool includeDeepContent = true)
	{
		return BuildResult(hostDocument, family, includeDeepContent).Fingerprint;
	}

	public static string Build(Document hostDocument, Family family, IDictionary<string, string> contentFingerprintCache, bool includeDeepContent = true)
	{
		if (contentFingerprintCache == null)
		{
			return LoadableFamilyContentSignatureService.Build(hostDocument, family, includeDeepContent);
		}
		string cacheKey = BuildCacheKey(hostDocument, family, includeDeepContent);
		if (string.IsNullOrWhiteSpace(cacheKey))
		{
			return LoadableFamilyContentSignatureService.Build(hostDocument, family, includeDeepContent);
		}
		string cached = null;
		if (contentFingerprintCache.TryGetValue(cacheKey, out cached))
		{
			return cached ?? string.Empty;
		}
		return contentFingerprintCache[cacheKey] = LoadableFamilyContentSignatureService.Build(hostDocument, family, includeDeepContent);
	}

	public static LoadableFamilyContentSignatureResult BuildResult(Document hostDocument, Family family, bool includeDeepContent = true)
	{
		return EnsureFailureReason(includeDeepContent ? BuildCoreResult(hostDocument, family) : BuildFastCoreResult(hostDocument, family), includeDeepContent ? "Precise" : "Fast", "Signature builder returned an empty fingerprint.", family);
	}

	public static string BuildFromOpenFamilyDocument(Family family, Document familyDocument)
	{
		return BuildResultFromOpenFamilyDocument(family, familyDocument).Fingerprint;
	}

	public static LoadableFamilyContentSignatureResult BuildResultFromOpenFamilyDocument(Family family, Document familyDocument)
	{
		LoadableFamilyContentSignatureResult BuildResultFromOpenFamilyDocument;
		if (family == null || familyDocument == null)
		{
			BuildResultFromOpenFamilyDocument = BuildFailureResult("Precise", "Family or open family document was empty.", family);
		}
		else
		{
			try
			{
				BuildResultFromOpenFamilyDocument = EnsureFailureReason(BuildCoreResultFromOpenFamilyDocument(family, familyDocument), "Precise", "Open family document signature returned an empty fingerprint.", family);
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				BuildResultFromOpenFamilyDocument = BuildFailureResult("Precise", "Open family document signature failed: " + ex2.Message, family);
				ProjectData.ClearProjectError();
			}
		}
		return BuildResultFromOpenFamilyDocument;
	}

	public static LoadableFamilyContentSignatureResult BuildDiagnosticFailureResult(string mode, string reason, Family family)
	{
		return BuildFailureResult(mode, reason, family);
	}

	private static LoadableFamilyContentSignatureResult BuildFastCoreResult(Document hostDocument, Family family)
	{
		LoadableFamilyContentSignatureResult BuildFastCoreResult;
		if (hostDocument == null || family == null)
		{
			BuildFastCoreResult = BuildFailureResult("Fast", "Host document or family was empty.", family);
		}
		else
		{
			try
			{
				List<string> typeParts = new List<string>();
				foreach (ElementId symbolId in family.GetFamilySymbolIds())
				{
					Element element = hostDocument.GetElement(symbolId);
					FamilySymbol symbol = (FamilySymbol)(object)((element is FamilySymbol) ? element : null);
					if (symbol != null)
					{
						typeParts.Add(Normalize(SafeElementName((Element)(object)symbol)) + "|" + Normalize(BuildElementParameterSignature(hostDocument, (Element)(object)symbol)));
					}
				}
				List<string> lines = new List<string>
				{
					"content-signature-version=fast-2",
					"family=" + Normalize(SafeElementName((Element)(object)family)),
					"category=" + Normalize(ResolveFamilyCategoryName(family)),
					"category-group=" + Normalize(ResolveFamilyCategoryGroup(family)),
					"shared=" + Normalize(ResolveIsShared(family).ToString()),
					"family-parameters=" + Normalize(BuildElementParameterSignature(hostDocument, (Element)(object)family)),
					"types=" + string.Join("\n", typeParts.OrderBy<string, string>([SpecialName] (string x) => x, StringComparer.Ordinal))
				};
				string signature = string.Join("\n", lines);
				BuildFastCoreResult = new LoadableFamilyContentSignatureResult
				{
					Fingerprint = HashString(signature),
					Signature = signature,
					Mode = "Fast"
				};
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				BuildFastCoreResult = BuildFailureResult("Fast", "Fast signature failed: " + ex2.Message, family);
				ProjectData.ClearProjectError();
			}
		}
		return BuildFastCoreResult;
	}

	private static LoadableFamilyContentSignatureResult BuildCoreResult(Document hostDocument, Family family)
	{
		LoadableFamilyContentSignatureResult BuildCoreResult;
		if (hostDocument == null || family == null)
		{
			BuildCoreResult = BuildFailureResult("Precise", "Host document or family was empty.", family);
		}
		else if (!CanEditFamilyDocument(family))
		{
			BuildCoreResult = BuildFailureResult("Precise", "Family is not editable or is in-place.", family);
		}
		else
		{
			Document familyDocument = null;
			try
			{
				familyDocument = hostDocument.EditFamily(family);
				BuildCoreResult = ((familyDocument != null) ? BuildCoreResultFromOpenFamilyDocument(family, familyDocument) : BuildFailureResult("Precise", "EditFamily returned no document.", family));
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				BuildCoreResult = BuildFailureResult("Precise", "EditFamily signature failed: " + ex2.GetType().Name + " - " + ex2.Message, family);
				ProjectData.ClearProjectError();
			}
			finally
			{
				if (familyDocument != null)
				{
					try
					{
						familyDocument.Close(false);
					}
					catch (Exception projectError)
					{
						ProjectData.SetProjectError(projectError);
						ProjectData.ClearProjectError();
					}
				}
			}
		}
		return BuildCoreResult;
	}

	private static LoadableFamilyContentSignatureResult BuildCoreResultFromOpenFamilyDocument(Family family, Document familyDocument)
	{
		if (family == null || familyDocument == null)
		{
			return BuildFailureResult("Precise", "Family or open family document was empty.", family);
		}
		bool includeAnnotationGraphics = IsAnnotationFamily(family);
		List<string> elementDebugLines = new List<string>();
		string elementSignature = BuildElementSignature(familyDocument, includeAnnotationGraphics, elementDebugLines);
		List<string> lines = new List<string>
		{
			"content-signature-version=6",
			"family=" + Normalize(SafeElementName((Element)(object)family)),
			"category=" + Normalize(ResolveFamilyCategoryName(family)),
			"elements=" + elementSignature,
			"types=" + BuildFamilyManagerSignature(familyDocument)
		};
		string signature = string.Join("\n", lines);
		return new LoadableFamilyContentSignatureResult
		{
			Fingerprint = HashString(signature),
			Signature = signature,
			DebugMetadata = BuildDebugMetadata(elementDebugLines),
			Mode = "Precise"
		};
	}

	private static LoadableFamilyContentSignatureResult BuildFailureResult(string mode, string reason, Family family)
	{
		string safeMode = (string.IsNullOrWhiteSpace(mode) ? "Unknown" : mode);
		List<string> lines = new List<string>
		{
			"content-signature-version=diagnostic-failure-1",
			"signature-mode=" + Normalize(safeMode),
			"signature-status=failed",
			"family=" + Normalize((family == null) ? string.Empty : SafeElementName((Element)(object)family)),
			"category=" + Normalize((family == null) ? string.Empty : ResolveFamilyCategoryName(family)),
			"category-group=" + Normalize((family == null) ? string.Empty : ResolveFamilyCategoryGroup(family)),
			"shared=" + Normalize((family == null) ? string.Empty : ResolveIsShared(family).ToString()),
			"error=" + Normalize(reason)
		};
		return new LoadableFamilyContentSignatureResult
		{
			Fingerprint = string.Empty,
			Signature = string.Join("\n", lines),
			Mode = safeMode,
			ErrorMessage = (reason ?? string.Empty)
		};
	}

	private static LoadableFamilyContentSignatureResult EnsureFailureReason(LoadableFamilyContentSignatureResult result, string mode, string fallbackReason, Family family)
	{
		if (result == null)
		{
			return BuildFailureResult(mode, fallbackReason + " No signature result was returned.", family);
		}
		if (!string.IsNullOrWhiteSpace(result.Fingerprint))
		{
			return result;
		}
		if (string.IsNullOrWhiteSpace(result.Mode))
		{
			result.Mode = (string.IsNullOrWhiteSpace(mode) ? "Unknown" : mode);
		}
		if (string.IsNullOrWhiteSpace(result.ErrorMessage))
		{
			result.ErrorMessage = (string.IsNullOrWhiteSpace(fallbackReason) ? "Signature builder returned an empty fingerprint." : fallbackReason);
		}
		if (string.IsNullOrWhiteSpace(result.Signature))
		{
			LoadableFamilyContentSignatureResult failure = BuildFailureResult(result.Mode, result.ErrorMessage, family);
			result.Signature = failure.Signature;
		}
		return result;
	}

	public static string BuildCacheKey(Document hostDocument, Family family, bool includeDeepContent)
	{
		if (hostDocument == null || family == null)
		{
			return string.Empty;
		}
		string documentKey = string.Empty;
		try
		{
			documentKey = hostDocument.PathName ?? string.Empty;
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			documentKey = string.Empty;
			ProjectData.ClearProjectError();
		}
		if (string.IsNullOrWhiteSpace(documentKey))
		{
			try
			{
				documentKey = hostDocument.Title ?? string.Empty;
			}
			catch (Exception projectError2)
			{
				ProjectData.SetProjectError(projectError2);
				documentKey = string.Empty;
				ProjectData.ClearProjectError();
			}
		}
		return (includeDeepContent ? "deep" : "fast") + "|" + Normalize(documentKey) + "|" + Normalize(((Element)family).UniqueId ?? string.Empty) + "|" + Normalize(SafeElementName((Element)(object)family)) + "|" + Normalize(ResolveFamilyCategoryName(family));
	}

	private static bool CanEditFamilyDocument(Family family)
	{
		bool CanEditFamilyDocument;
		try
		{
			CanEditFamilyDocument = family != null && !family.IsInPlace && family.IsEditable;
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			CanEditFamilyDocument = false;
			ProjectData.ClearProjectError();
		}
		return CanEditFamilyDocument;
	}

	private static string BuildElementSignature(Document familyDocument, bool includeAnnotationGraphics, IList<string> elementDebugLines = null)
	{
		//IL_001d: Unknown result type (might be due to invalid IL or missing references)
		_Closure_0024__14_002D0 arg = default(_Closure_0024__14_002D0);
		_Closure_0024__14_002D0 CS_0024_003C_003E8__locals2 = new _Closure_0024__14_002D0(arg);
		CS_0024_003C_003E8__locals2._0024VB_0024Local_includeAnnotationGraphics = includeAnnotationGraphics;
		if (familyDocument == null)
		{
			return string.Empty;
		}
		List<Element> elements = (from Element x in (IEnumerable)new FilteredElementCollector(familyDocument).WhereElementIsNotElementType()
			where ShouldCaptureElement(x, CS_0024_003C_003E8__locals2._0024VB_0024Local_includeAnnotationGraphics)
			select x).OrderBy<Element, string>([SpecialName] (Element x) => Normalize(((object)x).GetType().Name), StringComparer.Ordinal).ThenBy<Element, string>([SpecialName] (Element x) => Normalize(ResolveCategoryName(x)), StringComparer.Ordinal).ThenBy<Element, string>([SpecialName] (Element x) => Normalize(SafeElementName(x)), StringComparer.Ordinal)
			.ToList();
		List<string> parts = new List<string>();
		foreach (Element element in elements)
		{
			string signatureLine = Normalize(((object)element).GetType().Name) + "|" + Normalize(ResolveCategoryName(element)) + "|" + Normalize(SafeElementName(element)) + "|" + Normalize(BuildElementTypeReferenceSignature(familyDocument, element)) + "|" + Normalize(BuildElementParameterSignature(familyDocument, element));
			parts.Add(signatureLine);
			AddElementDebugLine(elementDebugLines, signatureLine, element);
		}
		return string.Join("\n", parts.OrderBy<string, string>([SpecialName] (string x) => x, StringComparer.Ordinal));
	}

	private static string BuildDebugMetadata(IEnumerable<string> elementDebugLines)
	{
		if (elementDebugLines == null)
		{
			return string.Empty;
		}
		List<string> lines = elementDebugLines.Where([SpecialName] (string x) => !string.IsNullOrWhiteSpace(x)).OrderBy<string, string>([SpecialName] (string x) => x, StringComparer.Ordinal).ToList();
		if (lines.Count == 0)
		{
			return string.Empty;
		}
		return string.Join("\n", lines);
	}

	private static void AddElementDebugLine(IList<string> target, string signatureLine, Element element)
	{
		if (target != null && element != null && !string.IsNullOrWhiteSpace(signatureLine))
		{
			target.Add("element\t" + Normalize(signatureLine) + "\t" + BuildElementDebugIdentity(element));
		}
	}

	private static string BuildElementDebugIdentity(Element element)
	{
		if (element == null)
		{
			return string.Empty;
		}
		List<string> parts = new List<string>();
		try
		{
			if (element.Id != null)
			{
				parts.Add("id=" + RevitElementIdCompat.CompatIntegerValue(element.Id).ToString(CultureInfo.InvariantCulture));
			}
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ProjectData.ClearProjectError();
		}
		try
		{
			if (!string.IsNullOrWhiteSpace(element.UniqueId))
			{
				parts.Add("uniqueId=" + element.UniqueId);
			}
		}
		catch (Exception projectError2)
		{
			ProjectData.SetProjectError(projectError2);
			ProjectData.ClearProjectError();
		}
		try
		{
			string name = SafeElementName(element);
			if (!string.IsNullOrWhiteSpace(name))
			{
				AddDebugTextToken(parts, "name", name);
			}
		}
		catch (Exception projectError3)
		{
			ProjectData.SetProjectError(projectError3);
			ProjectData.ClearProjectError();
		}
		try
		{
			string categoryName = ResolveCategoryName(element);
			if (!string.IsNullOrWhiteSpace(categoryName))
			{
				AddDebugTextToken(parts, "category", categoryName);
			}
		}
		catch (Exception projectError4)
		{
			ProjectData.SetProjectError(projectError4);
			ProjectData.ClearProjectError();
		}
		try
		{
			FamilyInstance instance = (FamilyInstance)(object)((element is FamilyInstance) ? element : null);
			if (instance != null && instance.Symbol != null)
			{
				string symbolName = SafeElementName((Element)(object)instance.Symbol);
				string familyName = string.Empty;
				try
				{
					if (instance.Symbol.Family != null)
					{
						familyName = SafeElementName((Element)(object)instance.Symbol.Family);
					}
				}
				catch (Exception projectError5)
				{
					ProjectData.SetProjectError(projectError5);
					ProjectData.ClearProjectError();
				}
				AddDebugTextToken(parts, "nestedFamily", familyName);
				AddDebugTextToken(parts, "nestedType", symbolName);
			}
		}
		catch (Exception projectError6)
		{
			ProjectData.SetProjectError(projectError6);
			ProjectData.ClearProjectError();
		}
		try
		{
			FamilySymbol symbol = (FamilySymbol)(object)((element is FamilySymbol) ? element : null);
			if (symbol != null)
			{
				string familyName2 = string.Empty;
				try
				{
					if (symbol.Family != null)
					{
						familyName2 = SafeElementName((Element)(object)symbol.Family);
					}
				}
				catch (Exception projectError7)
				{
					ProjectData.SetProjectError(projectError7);
					ProjectData.ClearProjectError();
				}
				AddDebugTextToken(parts, "nestedFamily", familyName2);
				AddDebugTextToken(parts, "nestedType", SafeElementName((Element)(object)symbol));
			}
		}
		catch (Exception projectError8)
		{
			ProjectData.SetProjectError(projectError8);
			ProjectData.ClearProjectError();
		}
		try
		{
			Family family = (Family)(object)((element is Family) ? element : null);
			if (family != null)
			{
				AddDebugTextToken(parts, "family", SafeElementName((Element)(object)family));
			}
		}
		catch (Exception projectError9)
		{
			ProjectData.SetProjectError(projectError9);
			ProjectData.ClearProjectError();
		}
		return string.Join(" ", parts);
	}

	private static void AddDebugTextToken(IList<string> parts, string key, string value)
	{
		if (parts != null && !string.IsNullOrWhiteSpace(key) && !string.IsNullOrWhiteSpace(value))
		{
			string text = value.Replace("\r\n", " ").Replace("\r", " ").Replace("\n", " ")
				.Replace("\t", " ")
				.Replace('"', '\'')
				.Trim();
			if (!string.IsNullOrWhiteSpace(text))
			{
				parts.Add(key + "=\"" + text + "\"");
			}
		}
	}

	private static bool ShouldCaptureElement(Element element, bool includeAnnotationGraphics)
	{
		if (element == null)
		{
			return false;
		}
		if (element is View || element is ViewSheet)
		{
			return false;
		}
		string categoryName = ResolveCategoryName(element);
		if (string.Equals(categoryName, "Views", StringComparison.OrdinalIgnoreCase) || string.Equals(categoryName, "Sheets", StringComparison.OrdinalIgnoreCase))
		{
			return false;
		}
		if (IsDocumentStyleOrEnvironmentElement(element))
		{
			return false;
		}
		if (IsUnidentifiableInternalElement(element))
		{
			return false;
		}
		if (IsCosmeticFamilyElement(element, includeAnnotationGraphics))
		{
			return false;
		}
		return true;
	}

	private static bool IsDocumentStyleOrEnvironmentElement(Element element)
	{
		if (element == null)
		{
			return false;
		}
		string className = Normalize(((object)element).GetType().Name);
		switch (className)
		{
		case "appearanceassetelement":
		case "areavolumesettings":
		case "assemblycodetable":
		case "defaultdividesettings":
		case "fillpatternelement":
		case "graphicsstyle":
		case "keynotetable":
		case "linepatternelement":
		case "loadcase":
		case "loadnature":
		case "parameterelement":
		case "sharedparameterelement":
		case "structuralsettings":
		case "sunandshadowsettings":
		case "viewnavigationtoolsettings":
			return true;
		default:
			if (string.Equals(className, "element", StringComparison.OrdinalIgnoreCase))
			{
				string categoryName = Normalize(ResolveCategoryName(element));
				if (string.Equals(categoryName, "cameras", StringComparison.OrdinalIgnoreCase) || string.Equals(categoryName, "work plane grid", StringComparison.OrdinalIgnoreCase))
				{
					return true;
				}
				switch (Normalize(SafeElementName(element)))
				{
				case "autojoin tracker element":
				case "extentelem":
				case "line weights":
				case "project browser":
				case "project units":
				case "project view":
				case "site settings":
				case "sun path":
					return true;
				}
			}
			return false;
		}
	}

	private static bool IsUnidentifiableInternalElement(Element element)
	{
		if (element == null)
		{
			return false;
		}
		if ((object)((object)element).GetType() != typeof(Element))
		{
			return false;
		}
		if (!string.IsNullOrWhiteSpace(ResolveCategoryName(element)))
		{
			return false;
		}
		if (!string.IsNullOrWhiteSpace(SafeElementName(element)))
		{
			return false;
		}
		try
		{
			ElementId typeId = element.GetTypeId();
			if (typeId != null && typeId != ElementId.InvalidElementId && RevitElementIdCompat.CompatIntegerValue(typeId) > 0)
			{
				return false;
			}
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ProjectData.ClearProjectError();
		}
		try
		{
			if (element.Parameters != null && ((IEnumerable)element.Parameters).Cast<Parameter>().Any([SpecialName] (Parameter x) => ShouldCaptureParameter(x)))
			{
				return false;
			}
		}
		catch (Exception projectError2)
		{
			ProjectData.SetProjectError(projectError2);
			ProjectData.ClearProjectError();
		}
		return true;
	}

	private static bool IsCosmeticFamilyElement(Element element, bool includeAnnotationGraphics)
	{
		//IL_0125: Unknown result type (might be due to invalid IL or missing references)
		//IL_012b: Invalid comparison between Unknown and I4
		if (element == null || includeAnnotationGraphics)
		{
			return false;
		}
		string className = Normalize(((object)element).GetType().Name);
		switch (className)
		{
		case "textnote":
		case "textelement":
		case "filledregion":
		case "detailcurve":
		case "detailline":
		case "detailarc":
		case "detailnurbsspline":
		case "detailellipse":
			return true;
		default:
			try
			{
				if (element.Category != null && (int)element.Category.CategoryType == 2)
				{
					if (string.Equals(className, "dimension", StringComparison.Ordinal))
					{
						return false;
					}
					return true;
				}
			}
			catch (Exception projectError)
			{
				ProjectData.SetProjectError(projectError);
				ProjectData.ClearProjectError();
			}
			return false;
		}
	}

	private static string BuildElementTypeReferenceSignature(Document familyDocument, Element element)
	{
		string BuildElementTypeReferenceSignature;
		try
		{
			ElementId typeId = element.GetTypeId();
			if (typeId == null || typeId == ElementId.InvalidElementId)
			{
				BuildElementTypeReferenceSignature = string.Empty;
			}
			else
			{
				Element elementType = familyDocument.GetElement(typeId);
				BuildElementTypeReferenceSignature = ((elementType != null) ? (((object)elementType).GetType().Name + "|" + ResolveCategoryName(elementType) + "|" + SafeElementName(elementType)) : RevitElementIdCompat.CompatIntegerValue(typeId).ToString(CultureInfo.InvariantCulture));
			}
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			BuildElementTypeReferenceSignature = string.Empty;
			ProjectData.ClearProjectError();
		}
		return BuildElementTypeReferenceSignature;
	}

	private static string BuildElementParameterSignature(Document familyDocument, Element element)
	{
		List<string> parts = new List<string>();
		try
		{
			using IEnumerator<Parameter> enumerator = (from Parameter x in (IEnumerable)element.Parameters
				where ShouldCaptureParameter(x)
				select x).OrderBy<Parameter, string>([SpecialName] (Parameter x) => Normalize(ResolveParameterName(x)), StringComparer.Ordinal).GetEnumerator();
			_Closure_0024__24_002D0 closure_0024__24_002D = default(_Closure_0024__24_002D0);
			while (enumerator.MoveNext())
			{
				closure_0024__24_002D = new _Closure_0024__24_002D0(closure_0024__24_002D);
				closure_0024__24_002D._0024VB_0024Local_parameter = enumerator.Current;
				parts.Add(Normalize(ResolveParameterName(closure_0024__24_002D._0024VB_0024Local_parameter)) + ":" + Normalize(ResolveStorageTypeName(closure_0024__24_002D._0024VB_0024Local_parameter)) + ":" + SafeBool(closure_0024__24_002D._Lambda_0024__2) + ":" + Normalize(BuildPortableParameterIdentity(closure_0024__24_002D._0024VB_0024Local_parameter)) + ":" + Normalize(BuildAssociatedFamilyParameterSignature(familyDocument, closure_0024__24_002D._0024VB_0024Local_parameter)));
			}
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ProjectData.ClearProjectError();
		}
		return string.Join("|", parts.OrderBy<string, string>([SpecialName] (string x) => x, StringComparer.Ordinal));
	}

	private static string BuildBoundingBoxSignature(Element element)
	{
		string BuildBoundingBoxSignature;
		try
		{
			BoundingBoxXYZ box = element[(View)null];
			BuildBoundingBoxSignature = ((box != null) ? ("min=" + FormatXyz(box.Min) + ";max=" + FormatXyz(box.Max)) : string.Empty);
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			BuildBoundingBoxSignature = string.Empty;
			ProjectData.ClearProjectError();
		}
		return BuildBoundingBoxSignature;
	}

	private static string BuildGeometrySignature(Element element, Options options)
	{
		List<string> parts = new List<string>();
		try
		{
			GeometryElement geometry = element[options];
			if (geometry != null)
			{
				AppendGeometrySignatures(geometry, parts);
			}
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ProjectData.ClearProjectError();
		}
		return string.Join("|", parts.OrderBy<string, string>([SpecialName] (string x) => x, StringComparer.Ordinal));
	}

	private static void AppendGeometrySignatures(GeometryElement geometry, ICollection<string> parts)
	{
		if (geometry == null || parts == null)
		{
			return;
		}
		foreach (GeometryObject item in geometry)
		{
			AppendGeometryObjectSignature(item, parts);
		}
	}

	private static void AppendGeometryObjectSignature(GeometryObject geometryObject, ICollection<string> parts)
	{
		//IL_0018: Unknown result type (might be due to invalid IL or missing references)
		//IL_001e: Expected O, but got Unknown
		//IL_00c1: Unknown result type (might be due to invalid IL or missing references)
		//IL_00c7: Expected O, but got Unknown
		//IL_00f4: Unknown result type (might be due to invalid IL or missing references)
		//IL_00fa: Expected O, but got Unknown
		//IL_0152: Unknown result type (might be due to invalid IL or missing references)
		//IL_0159: Expected O, but got Unknown
		//IL_017d: Unknown result type (might be due to invalid IL or missing references)
		//IL_0184: Expected O, but got Unknown
		if (geometryObject == null || parts == null)
		{
			return;
		}
		if (geometryObject is Solid)
		{
			Solid solid = (Solid)geometryObject;
			if (solid.Faces != null && solid.Faces.Size != 0)
			{
				parts.Add("solid:faces=" + solid.Faces.Size.ToString(CultureInfo.InvariantCulture) + ";edges=" + SafeEdgeCount(solid).ToString(CultureInfo.InvariantCulture) + ";volume=" + FormatDouble(SafeVolume(solid)) + ";area=" + FormatDouble(SafeSurfaceArea(solid)));
			}
			return;
		}
		if (geometryObject is Mesh)
		{
			Mesh mesh = (Mesh)geometryObject;
			parts.Add("mesh:triangles=" + mesh.NumTriangles.ToString(CultureInfo.InvariantCulture));
			return;
		}
		if (geometryObject is Curve)
		{
			Curve curve = (Curve)geometryObject;
			parts.Add("curve:" + ((object)curve).GetType().Name + ";length=" + FormatDouble(SafeCurveLength(curve)) + ";ends=" + BuildCurveEndpointSignature(curve));
			return;
		}
		if (geometryObject is GeometryInstance)
		{
			GeometryInstance instance = (GeometryInstance)geometryObject;
			try
			{
				AppendGeometrySignatures(instance.GetInstanceGeometry(), parts);
				return;
			}
			catch (Exception projectError)
			{
				ProjectData.SetProjectError(projectError);
				ProjectData.ClearProjectError();
				return;
			}
		}
		if (geometryObject is Point)
		{
			Point point = (Point)geometryObject;
			parts.Add("point:" + FormatXyz(point.Coord));
		}
		else
		{
			parts.Add("geometry:" + ((object)geometryObject).GetType().Name);
		}
	}

	private static string BuildCurveEndpointSignature(Curve curve)
	{
		string BuildCurveEndpointSignature;
		if (curve == null)
		{
			BuildCurveEndpointSignature = string.Empty;
		}
		else
		{
			try
			{
				BuildCurveEndpointSignature = (curve.IsBound ? (FormatXyz(curve.GetEndPoint(0)) + ">" + FormatXyz(curve.GetEndPoint(1))) : "unbound");
			}
			catch (Exception projectError)
			{
				ProjectData.SetProjectError(projectError);
				BuildCurveEndpointSignature = string.Empty;
				ProjectData.ClearProjectError();
			}
		}
		return BuildCurveEndpointSignature;
	}

	private static string BuildFamilyManagerSignature(Document familyDocument)
	{
		//IL_0031: Unknown result type (might be due to invalid IL or missing references)
		//IL_0038: Expected O, but got Unknown
		//IL_008f: Unknown result type (might be due to invalid IL or missing references)
		//IL_0096: Expected O, but got Unknown
		string BuildFamilyManagerSignature;
		try
		{
			FamilyManager manager = familyDocument.FamilyManager;
			if (manager == null)
			{
				BuildFamilyManagerSignature = string.Empty;
			}
			else
			{
				List<string> parameterParts = new List<string>();
				foreach (FamilyParameter parameter in manager.Parameters)
				{
					FamilyParameter familyParameter = parameter;
					if (familyParameter != null && familyParameter.Definition != null)
					{
						parameterParts.Add(BuildFamilyParameterDefinitionSignature(familyParameter));
					}
				}
				List<string> typeParts = new List<string>();
				foreach (FamilyType type in manager.Types)
				{
					FamilyType familyType = type;
					if (familyType != null)
					{
						typeParts.Add(Normalize(familyType.Name));
					}
				}
				BuildFamilyManagerSignature = "params=" + string.Join(";", parameterParts.OrderBy<string, string>([SpecialName] (string x) => x, StringComparer.Ordinal)) + "\ntypes=" + string.Join(";", typeParts.OrderBy<string, string>([SpecialName] (string x) => x, StringComparer.Ordinal)) + "\nnested-labels=" + BuildNestedFamilyLabelSignature(familyDocument, manager);
			}
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			BuildFamilyManagerSignature = string.Empty;
			ProjectData.ClearProjectError();
		}
		return BuildFamilyManagerSignature;
	}

	private static string BuildFamilyParameterDefinitionSignature(FamilyParameter familyParameter)
	{
		//IL_005c: Unknown result type (might be due to invalid IL or missing references)
		//IL_0061: Unknown result type (might be due to invalid IL or missing references)
		if (familyParameter == null || familyParameter.Definition == null)
		{
			return string.Empty;
		}
		return Normalize(familyParameter.Definition.Name) + "|" + Normalize(((Enum)familyParameter.StorageType/*cast due to .constrained prefix*/).ToString()) + "|" + Normalize(SafeBool([SpecialName] () => familyParameter.IsInstance).ToString()) + "|" + Normalize(ResolveFamilyParameterFormula(familyParameter)) + "|" + Normalize(ResolveFamilyParameterGuid(familyParameter));
	}

	private static string BuildNestedFamilyLabelSignature(Document familyDocument, FamilyManager manager)
	{
		//IL_0018: Unknown result type (might be due to invalid IL or missing references)
		if (familyDocument == null || manager == null)
		{
			return string.Empty;
		}
		List<string> parts = new List<string>();
		try
		{
			List<FamilyInstance> instances = ((IEnumerable)new FilteredElementCollector(familyDocument).OfClass(typeof(FamilyInstance))).Cast<FamilyInstance>().OrderBy<FamilyInstance, string>([SpecialName] (FamilyInstance x) => Normalize(ResolveCategoryName((Element)(object)x)), StringComparer.Ordinal).ThenBy<FamilyInstance, string>([SpecialName] (FamilyInstance x) => Normalize(SafeElementName((Element)(object)x)), StringComparer.Ordinal)
				.ToList();
			foreach (FamilyInstance instance in instances)
			{
				string instanceLabel = BuildFamilyInstanceReferenceSignature(instance);
				foreach (Parameter parameter in (from Parameter x in (IEnumerable)((Element)instance).Parameters
					where ShouldCaptureParameter(x)
					select x).OrderBy<Parameter, string>([SpecialName] (Parameter x) => Normalize(ResolveParameterName(x)), StringComparer.Ordinal))
				{
					FamilyParameter associated = ResolveAssociatedFamilyParameter(manager, parameter);
					if (associated != null)
					{
						string nestedTypeValue = BuildNestedFamilyTypeParameterValueSignature(familyDocument, parameter);
						if (!string.IsNullOrWhiteSpace(nestedTypeValue))
						{
							parts.Add(instanceLabel + "|" + Normalize(ResolveParameterName(parameter)) + "=>label=" + Normalize(ResolveFamilyParameterName(associated)) + "|" + nestedTypeValue);
						}
					}
				}
			}
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ProjectData.ClearProjectError();
		}
		return string.Join(";", parts.Distinct<string>(StringComparer.Ordinal).OrderBy<string, string>([SpecialName] (string x) => x, StringComparer.Ordinal));
	}

	private static string BuildNestedFamilyTypeParameterValueSignature(Document familyDocument, Parameter parameter)
	{
		//IL_0013: Unknown result type (might be due to invalid IL or missing references)
		//IL_0019: Invalid comparison between Unknown and I4
		string BuildNestedFamilyTypeParameterValueSignature;
		if (familyDocument == null || parameter == null)
		{
			BuildNestedFamilyTypeParameterValueSignature = string.Empty;
		}
		else
		{
			try
			{
				if ((int)parameter.StorageType != 4)
				{
					BuildNestedFamilyTypeParameterValueSignature = string.Empty;
				}
				else
				{
					ElementId valueId = parameter.AsElementId();
					if (valueId == null || valueId == ElementId.InvalidElementId)
					{
						BuildNestedFamilyTypeParameterValueSignature = "nested-family=|nested-type=|nested-category=";
					}
					else
					{
						Element valueElement = familyDocument.GetElement(valueId);
						if (valueElement == null)
						{
							BuildNestedFamilyTypeParameterValueSignature = "nested-family=|nested-type=" + Normalize(RevitElementIdCompat.CompatIntegerValue(valueId).ToString(CultureInfo.InvariantCulture)) + "|nested-category=";
						}
						else
						{
							FamilySymbol symbol = (FamilySymbol)(object)((valueElement is FamilySymbol) ? valueElement : null);
							if (symbol != null)
							{
								string familyName = string.Empty;
								try
								{
									if (symbol.Family != null)
									{
										familyName = SafeElementName((Element)(object)symbol.Family);
									}
								}
								catch (Exception projectError)
								{
									ProjectData.SetProjectError(projectError);
									ProjectData.ClearProjectError();
								}
								BuildNestedFamilyTypeParameterValueSignature = "nested-family=" + Normalize(familyName) + "|nested-type=" + Normalize(SafeElementName((Element)(object)symbol)) + "|nested-category=" + Normalize(ResolveCategoryName((Element)(object)symbol));
							}
							else
							{
								BuildNestedFamilyTypeParameterValueSignature = "nested-family=|nested-type=" + Normalize(SafeElementName(valueElement)) + "|nested-category=" + Normalize(ResolveCategoryName(valueElement));
							}
						}
					}
				}
			}
			catch (Exception projectError2)
			{
				ProjectData.SetProjectError(projectError2);
				BuildNestedFamilyTypeParameterValueSignature = string.Empty;
				ProjectData.ClearProjectError();
			}
		}
		return BuildNestedFamilyTypeParameterValueSignature;
	}

	private static string BuildFamilyInstanceReferenceSignature(FamilyInstance instance)
	{
		if (instance == null)
		{
			return string.Empty;
		}
		string symbolName = string.Empty;
		string familyName = string.Empty;
		try
		{
			if (instance.Symbol != null)
			{
				symbolName = SafeElementName((Element)(object)instance.Symbol);
				if (instance.Symbol.Family != null)
				{
					familyName = SafeElementName((Element)(object)instance.Symbol.Family);
				}
			}
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ProjectData.ClearProjectError();
		}
		return Normalize(ResolveCategoryName((Element)(object)instance)) + "|" + Normalize(familyName) + "|" + Normalize(symbolName) + "|" + Normalize(SafeElementName((Element)(object)instance));
	}

	private static string BuildAssociatedFamilyParameterSignature(Document familyDocument, Parameter parameter)
	{
		string BuildAssociatedFamilyParameterSignature;
		if (familyDocument == null || parameter == null)
		{
			BuildAssociatedFamilyParameterSignature = string.Empty;
		}
		else
		{
			try
			{
				FamilyParameter associated = ResolveAssociatedFamilyParameter(familyDocument.FamilyManager, parameter);
				BuildAssociatedFamilyParameterSignature = ((associated != null) ? BuildFamilyParameterDefinitionSignature(associated) : string.Empty);
			}
			catch (Exception projectError)
			{
				ProjectData.SetProjectError(projectError);
				BuildAssociatedFamilyParameterSignature = string.Empty;
				ProjectData.ClearProjectError();
			}
		}
		return BuildAssociatedFamilyParameterSignature;
	}

	private static FamilyParameter ResolveAssociatedFamilyParameter(FamilyManager manager, Parameter parameter)
	{
		FamilyParameter ResolveAssociatedFamilyParameter;
		if (manager == null || parameter == null)
		{
			ResolveAssociatedFamilyParameter = null;
		}
		else
		{
			try
			{
				MethodInfo methodInfo = ((object)manager).GetType().GetMethod("GetAssociatedFamilyParameter", new Type[1] { typeof(Parameter) });
				if ((object)methodInfo == null)
				{
					ResolveAssociatedFamilyParameter = null;
				}
				else
				{
					object? obj = methodInfo.Invoke(manager, new object[1] { parameter });
					ResolveAssociatedFamilyParameter = (FamilyParameter)((obj is FamilyParameter) ? obj : null);
				}
			}
			catch (Exception projectError)
			{
				ProjectData.SetProjectError(projectError);
				ResolveAssociatedFamilyParameter = null;
				ProjectData.ClearProjectError();
			}
		}
		return ResolveAssociatedFamilyParameter;
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

	private static bool IsAnnotationFamily(Family family)
	{
		//IL_0011: Unknown result type (might be due to invalid IL or missing references)
		//IL_0017: Invalid comparison between Unknown and I4
		bool IsAnnotationFamily;
		try
		{
			IsAnnotationFamily = family != null && family.FamilyCategory != null && (int)family.FamilyCategory.CategoryType == 2;
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			IsAnnotationFamily = false;
			ProjectData.ClearProjectError();
		}
		return IsAnnotationFamily;
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

	private static string BuildPortableParameterIdentity(Parameter parameter)
	{
		if (parameter == null)
		{
			return string.Empty;
		}
		string externalGuid = ResolveExternalGuid(parameter);
		if (!string.IsNullOrWhiteSpace(externalGuid))
		{
			return "guid:" + externalGuid;
		}
		try
		{
			if (parameter.Id != null && RevitElementIdCompat.CompatIntegerValue(parameter.Id) < 0)
			{
				return "builtin:" + RevitElementIdCompat.CompatIntegerValue(parameter.Id).ToString(CultureInfo.InvariantCulture);
			}
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ProjectData.ClearProjectError();
		}
		return string.Empty;
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

	private static string ResolveFamilyParameterGuid(FamilyParameter familyParameter)
	{
		try
		{
			Definition definition = familyParameter.Definition;
			ExternalDefinition externalDefinition = (ExternalDefinition)(object)((definition is ExternalDefinition) ? definition : null);
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

	private static string ResolveFamilyCategoryName(Family family)
	{
		return FamilyBrowserFamilyClassificationService.ResolveCategoryName(family);
	}

	private static string ResolveFamilyCategoryGroup(Family family)
	{
		return FamilyBrowserFamilyClassificationService.ResolveCategoryGroup(family);
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

	private static string SafeElementName(Element element)
	{
		string SafeElementName;
		if (element == null)
		{
			SafeElementName = string.Empty;
		}
		else
		{
			try
			{
				SafeElementName = element.Name ?? string.Empty;
			}
			catch (Exception projectError)
			{
				ProjectData.SetProjectError(projectError);
				SafeElementName = string.Empty;
				ProjectData.ClearProjectError();
			}
		}
		return SafeElementName;
	}

	private static int SafeEdgeCount(Solid solid)
	{
		try
		{
			if (solid != null && solid.Edges != null)
			{
				return solid.Edges.Size;
			}
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ProjectData.ClearProjectError();
		}
		return 0;
	}

	private static double SafeVolume(Solid solid)
	{
		double SafeVolume;
		try
		{
			SafeVolume = solid.Volume;
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			SafeVolume = 0.0;
			ProjectData.ClearProjectError();
		}
		return SafeVolume;
	}

	private static double SafeSurfaceArea(Solid solid)
	{
		double SafeSurfaceArea;
		try
		{
			SafeSurfaceArea = solid.SurfaceArea;
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			SafeSurfaceArea = 0.0;
			ProjectData.ClearProjectError();
		}
		return SafeSurfaceArea;
	}

	private static double SafeCurveLength(Curve curve)
	{
		double SafeCurveLength;
		try
		{
			SafeCurveLength = curve.Length;
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			SafeCurveLength = 0.0;
			ProjectData.ClearProjectError();
		}
		return SafeCurveLength;
	}

	private static string FormatXyz(XYZ value)
	{
		if (value == null)
		{
			return string.Empty;
		}
		return FormatDouble(value.X) + "," + FormatDouble(value.Y) + "," + FormatDouble(value.Z);
	}

	private static string FormatDouble(double value)
	{
		return Math.Round(value, 9).ToString("G17", CultureInfo.InvariantCulture);
	}

	private static string HashString(string value)
	{
		using SHA256 sha = SHA256.Create();
		byte[] bytes = Encoding.UTF8.GetBytes(value ?? string.Empty);
		return BitConverter.ToString(sha.ComputeHash(bytes)).Replace("-", string.Empty).ToLowerInvariant();
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
		if (value == null)
		{
			return string.Empty;
		}
		return value.Trim().Replace("\r\n", "\n").Replace("\r", "\n")
			.ToLowerInvariant();
	}
}
