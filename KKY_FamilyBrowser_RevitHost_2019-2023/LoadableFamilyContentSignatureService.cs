using System;
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
					if (hostDocument.GetElement(symbolId) is FamilySymbol symbol)
					{
						typeParts.Add(Normalize(SafeElementName(symbol)) + "|" + Normalize(BuildElementParameterSignature(hostDocument, symbol)));
					}
				}
				List<string> lines = new List<string>
				{
					"content-signature-version=fast-2",
					"family=" + Normalize(SafeElementName(family)),
					"category=" + Normalize(ResolveFamilyCategoryName(family)),
					"category-group=" + Normalize(ResolveFamilyCategoryGroup(family)),
					"shared=" + Normalize(ResolveIsShared(family).ToString()),
					"family-parameters=" + Normalize(BuildElementParameterSignature(hostDocument, family)),
					"types=" + string.Join("\n", typeParts.OrderBy([SpecialName] (string x) => x, StringComparer.Ordinal))
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
						familyDocument.Close(saveModified: false);
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
		List<string> lookupDisplayLines = new List<string>();
		string lookupTableSignature = BuildLookupTableSignature(familyDocument, lookupDisplayLines);
		List<string> lines = new List<string>
		{
			"content-signature-version=7",
			"family=" + Normalize(SafeElementName(family)),
			"category=" + Normalize(ResolveFamilyCategoryName(family)),
			"elements=" + elementSignature,
			"types=" + BuildFamilyManagerSignature(familyDocument),
			"lookup-tables=" + lookupTableSignature
		};
		string signature = string.Join("\n", lines);
		return new LoadableFamilyContentSignatureResult
		{
			Fingerprint = HashString(signature),
			Signature = signature,
			DebugMetadata = BuildDebugMetadata(elementDebugLines.Concat(lookupDisplayLines)),
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
			"family=" + Normalize((family == null) ? string.Empty : SafeElementName(family)),
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
		return (includeDeepContent ? "deep" : "fast") + "|" + Normalize(documentKey) + "|" + Normalize(family.UniqueId ?? string.Empty) + "|" + Normalize(SafeElementName(family)) + "|" + Normalize(ResolveFamilyCategoryName(family));
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
		_Closure_0024__14_002D0 arg = default(_Closure_0024__14_002D0);
		_Closure_0024__14_002D0 CS_0024_003C_003E8__locals2 = new _Closure_0024__14_002D0(arg);
		CS_0024_003C_003E8__locals2._0024VB_0024Local_includeAnnotationGraphics = includeAnnotationGraphics;
		if (familyDocument == null)
		{
			return string.Empty;
		}
		List<Element> elements = (from Element x in new FilteredElementCollector(familyDocument).WhereElementIsNotElementType()
			where ShouldCaptureElement(x, CS_0024_003C_003E8__locals2._0024VB_0024Local_includeAnnotationGraphics)
			select x).OrderBy([SpecialName] (Element x) => Normalize(x.GetType().Name), StringComparer.Ordinal).ThenBy([SpecialName] (Element x) => Normalize(ResolveCategoryName(x)), StringComparer.Ordinal).ThenBy([SpecialName] (Element x) => Normalize(SafeElementName(x)), StringComparer.Ordinal)
			.ToList();
		List<string> parts = new List<string>();
		foreach (Element element in elements)
		{
			string signatureLine = Normalize(element.GetType().Name) + "|" + Normalize(ResolveCategoryName(element)) + "|" + Normalize(SafeElementName(element)) + "|" + Normalize(BuildElementTypeReferenceSignature(familyDocument, element)) + "|" + Normalize(BuildElementParameterSignature(familyDocument, element));
			parts.Add(signatureLine);
			AddElementDebugLine(elementDebugLines, signatureLine, element);
		}
		return string.Join("\n", parts.OrderBy([SpecialName] (string x) => x, StringComparer.Ordinal));
	}

	private static string BuildDebugMetadata(IEnumerable<string> elementDebugLines)
	{
		if (elementDebugLines == null)
		{
			return string.Empty;
		}
		List<string> lines = elementDebugLines.Where([SpecialName] (string x) => !string.IsNullOrWhiteSpace(x)).OrderBy([SpecialName] (string x) => x, StringComparer.Ordinal).ToList();
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
			if ((object)element.Id != null)
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
			if (element is FamilyInstance { Symbol: not null } instance)
			{
				string symbolName = SafeElementName(instance.Symbol);
				string familyName = string.Empty;
				try
				{
					if (instance.Symbol.Family != null)
					{
						familyName = SafeElementName(instance.Symbol.Family);
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
			if (element is FamilySymbol symbol)
			{
				string familyName2 = string.Empty;
				try
				{
					if (symbol.Family != null)
					{
						familyName2 = SafeElementName(symbol.Family);
					}
				}
				catch (Exception projectError7)
				{
					ProjectData.SetProjectError(projectError7);
					ProjectData.ClearProjectError();
				}
				AddDebugTextToken(parts, "nestedFamily", familyName2);
				AddDebugTextToken(parts, "nestedType", SafeElementName(symbol));
			}
		}
		catch (Exception projectError8)
		{
			ProjectData.SetProjectError(projectError8);
			ProjectData.ClearProjectError();
		}
		try
		{
			if (element is Family family)
			{
				AddDebugTextToken(parts, "family", SafeElementName(family));
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
		string className = Normalize(element.GetType().Name);
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
		if ((object)element.GetType() != typeof(Element))
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
			if ((object)typeId != null && typeId != ElementId.InvalidElementId && RevitElementIdCompat.CompatIntegerValue(typeId) > 0)
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
			if (element.Parameters != null && element.Parameters.Cast<Parameter>().Any([SpecialName] (Parameter x) => ShouldCaptureParameter(x)))
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
		if (element == null || includeAnnotationGraphics)
		{
			return false;
		}
		string className = Normalize(element.GetType().Name);
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
				if (element.Category != null && element.Category.CategoryType == CategoryType.Annotation)
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
			if ((object)typeId == null || typeId == ElementId.InvalidElementId)
			{
				BuildElementTypeReferenceSignature = string.Empty;
			}
			else
			{
				Element elementType = familyDocument.GetElement(typeId);
				BuildElementTypeReferenceSignature = ((elementType != null) ? (elementType.GetType().Name + "|" + ResolveCategoryName(elementType) + "|" + SafeElementName(elementType)) : RevitElementIdCompat.CompatIntegerValue(typeId).ToString(CultureInfo.InvariantCulture));
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
			using IEnumerator<Parameter> enumerator = (from Parameter x in element.Parameters
				where ShouldCaptureParameter(x)
				select x).OrderBy([SpecialName] (Parameter x) => Normalize(ResolveParameterName(x)), StringComparer.Ordinal).GetEnumerator();
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
		return string.Join("|", parts.OrderBy([SpecialName] (string x) => x, StringComparer.Ordinal));
	}

	private static string BuildBoundingBoxSignature(Element element)
	{
		string BuildBoundingBoxSignature;
		try
		{
			BoundingBoxXYZ box = element.get_BoundingBox((View)null);
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
			GeometryElement geometry = element.get_Geometry(options);
			if ((object)geometry != null)
			{
				AppendGeometrySignatures(geometry, parts);
			}
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ProjectData.ClearProjectError();
		}
		return string.Join("|", parts.OrderBy([SpecialName] (string x) => x, StringComparer.Ordinal));
	}

	private static void AppendGeometrySignatures(GeometryElement geometry, ICollection<string> parts)
	{
		if ((object)geometry == null || parts == null)
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
		if ((object)geometryObject == null || parts == null)
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
			parts.Add("curve:" + curve.GetType().Name + ";length=" + FormatDouble(SafeCurveLength(curve)) + ";ends=" + BuildCurveEndpointSignature(curve));
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
			parts.Add("geometry:" + geometryObject.GetType().Name);
		}
	}

	private static string BuildCurveEndpointSignature(Curve curve)
	{
		string BuildCurveEndpointSignature;
		if ((object)curve == null)
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
				foreach (FamilyParameter familyParameter in manager.Parameters)
				{
					if (familyParameter != null && familyParameter.Definition != null)
					{
						parameterParts.Add(BuildFamilyParameterDefinitionSignature(familyParameter));
					}
				}
				List<string> typeParts = new List<string>();
				foreach (FamilyType familyType in manager.Types)
				{
					if (familyType != null)
					{
						typeParts.Add(Normalize(familyType.Name));
					}
				}
				BuildFamilyManagerSignature = "params=" + string.Join(";", parameterParts.OrderBy([SpecialName] (string x) => x, StringComparer.Ordinal)) + "\ntypes=" + string.Join(";", typeParts.OrderBy([SpecialName] (string x) => x, StringComparer.Ordinal)) + "\nnested-labels=" + BuildNestedFamilyLabelSignature(familyDocument, manager);
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

	private static string BuildLookupTableSignature(Document familyDocument, IList<string> displayLines)
	{
		if (familyDocument == null)
		{
			return string.Empty;
		}
		List<string> parts = new List<string>();
		object manager = null;
		try
		{
			manager = ResolveFamilySizeTableManager(familyDocument);
			if (manager == null)
			{
				return string.Empty;
			}
			foreach (string tableName in ReadFamilySizeTableNames(manager).Where((string x) => !string.IsNullOrWhiteSpace(x)).Distinct(StringComparer.OrdinalIgnoreCase).OrderBy((string x) => Normalize(x), StringComparer.Ordinal))
			{
				object table = null;
				try
				{
					table = InvokePublicInstanceMethod(manager, "GetSizeTable", tableName);
					if (table == null)
					{
						parts.Add("table=" + NormalizeLookupToken(tableName) + "|missing");
						displayLines?.Add("lookup-display-table=" + NormalizeLookupDisplayToken(tableName) + "|missing");
					}
					else
					{
						int rowCount = ReadIntProperty(table, "NumberOfRows");
						int columnCount = ReadIntProperty(table, "NumberOfColumns");
						int columnOffset = ResolveFamilySizeTableColumnOffset(table, columnCount);
						parts.Add("table=" + NormalizeLookupToken(tableName) + "|columns=" + BuildFamilySizeTableColumnSignature(table, columnCount, columnOffset) + "|rows=" + BuildFamilySizeTableRowSignature(table, rowCount, columnCount, columnOffset));
						displayLines?.Add("lookup-display-table=" + NormalizeLookupDisplayToken(tableName) + "|columns=" + columnCount.ToString(CultureInfo.InvariantCulture) + "|rows=" + rowCount.ToString(CultureInfo.InvariantCulture));
					}
				}
				catch (Exception ex)
				{
					ProjectData.SetProjectError(ex);
					parts.Add("table=" + NormalizeLookupToken(tableName) + "|error=" + NormalizeLookupToken(ex.GetType().Name + ":" + ex.Message));
					displayLines?.Add("lookup-display-table=" + NormalizeLookupDisplayToken(tableName) + "|error=" + NormalizeLookupDisplayToken(ex.GetType().Name));
					ProjectData.ClearProjectError();
				}
				finally
				{
					DisposeIfNeeded(table);
				}
			}
		}
		catch (Exception ex2)
		{
			ProjectData.SetProjectError(ex2);
			parts.Add("lookup-table-error=" + NormalizeLookupToken(ex2.GetType().Name + ":" + ex2.Message));
			displayLines?.Add("lookup-display-table=lookup-table-error|error=" + NormalizeLookupDisplayToken(ex2.GetType().Name));
			ProjectData.ClearProjectError();
		}
		finally
		{
			DisposeIfNeeded(manager);
		}
		return string.Join("\n", parts.OrderBy((string x) => x, StringComparer.Ordinal));
	}

	private static object ResolveFamilySizeTableManager(Document familyDocument)
	{
		if (familyDocument == null)
		{
			return null;
		}
		try
		{
			Type managerType = typeof(FamilySizeTableManager);
			List<MethodInfo> methods = managerType.GetMethods(BindingFlags.Public | BindingFlags.Static).Where((MethodInfo x) => string.Equals(x.Name, "GetFamilySizeTableManager", StringComparison.Ordinal)).ToList();
			foreach (MethodInfo method in methods)
			{
				ParameterInfo[] parameters = method.GetParameters();
				if (parameters.Length == 1 && parameters[0].ParameterType.IsAssignableFrom(typeof(Document)))
				{
					return method.Invoke(null, new object[1] { familyDocument });
				}
			}
			ElementId ownerFamilyId = ResolveFamilySizeTableOwnerFamilyId(familyDocument);
			if ((object)ownerFamilyId == null || ownerFamilyId == ElementId.InvalidElementId)
			{
				return null;
			}
			foreach (MethodInfo method2 in methods)
			{
				ParameterInfo[] parameters2 = method2.GetParameters();
				if (parameters2.Length == 2 && parameters2[0].ParameterType.IsAssignableFrom(typeof(Document)) && parameters2[1].ParameterType.IsAssignableFrom(typeof(ElementId)))
				{
					return method2.Invoke(null, new object[2] { familyDocument, ownerFamilyId });
				}
			}
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ProjectData.ClearProjectError();
		}
		return null;
	}

	private static ElementId ResolveFamilySizeTableOwnerFamilyId(Document familyDocument)
	{
		if (familyDocument == null)
		{
			return ElementId.InvalidElementId;
		}
		try
		{
			Family ownerFamily = familyDocument.OwnerFamily;
			if (ownerFamily != null && (object)ownerFamily.Id != null && ownerFamily.Id != ElementId.InvalidElementId)
			{
				return ownerFamily.Id;
			}
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ProjectData.ClearProjectError();
		}
		return ElementId.InvalidElementId;
	}

	private static List<string> ReadFamilySizeTableNames(object manager)
	{
		List<string> result = new List<string>();
		if (manager == null)
		{
			return result;
		}
		try
		{
			object names = InvokePublicInstanceMethod(manager, "GetAllSizeTableNames");
			if (names is System.Collections.IEnumerable enumerable && !(names is string))
			{
				foreach (object item in enumerable)
				{
					string text = Convert.ToString(item, CultureInfo.InvariantCulture);
					if (!string.IsNullOrWhiteSpace(text))
					{
						result.Add(text.Trim());
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

	private static int ResolveFamilySizeTableColumnOffset(object table, int columnCount)
	{
		int successZero = 0;
		int successOne = 0;
		ReadFamilySizeTableColumns(table, columnCount, 0, out successZero);
		ReadFamilySizeTableColumns(table, columnCount, 1, out successOne);
		return (successOne > successZero) ? 1 : 0;
	}

	private static string BuildFamilySizeTableColumnSignature(object table, int columnCount, int columnOffset)
	{
		int successCount = 0;
		return string.Join(";", ReadFamilySizeTableColumns(table, columnCount, columnOffset, out successCount));
	}

	private static List<string> ReadFamilySizeTableColumns(object table, int columnCount, int columnOffset, out int successCount)
	{
		List<string> result = new List<string>();
		successCount = 0;
		if (table == null || columnCount <= 0)
		{
			return result;
		}
		for (int columnIndex = 0; columnIndex < columnCount; columnIndex++)
		{
			object header = null;
			try
			{
				header = InvokePublicInstanceMethod(table, "GetColumnHeader", checked(columnIndex + columnOffset));
				if (header == null)
				{
					result.Add(columnIndex.ToString(CultureInfo.InvariantCulture) + ":missing");
				}
				else
				{
					successCount = checked(successCount + 1);
					result.Add(columnIndex.ToString(CultureInfo.InvariantCulture) + ":" + NormalizeLookupToken(ReadStringProperty(header, "Name")) + ":" + NormalizeLookupToken(ReadStringProperty(header, "UnitType")) + ":" + NormalizeLookupToken(ReadStringProperty(header, "DisplayUnitType")));
				}
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				result.Add(columnIndex.ToString(CultureInfo.InvariantCulture) + ":error=" + NormalizeLookupToken(ex.GetType().Name));
				ProjectData.ClearProjectError();
			}
			finally
			{
				DisposeIfNeeded(header);
			}
		}
		return result;
	}

	private static string BuildFamilySizeTableRowSignature(object table, int rowCount, int columnCount, int columnOffset)
	{
		int successZero = 0;
		int successOne = 0;
		List<string> zeroBasedRows = ReadFamilySizeTableRows(table, rowCount, columnCount, 0, columnOffset, out successZero);
		List<string> oneBasedRows = ReadFamilySizeTableRows(table, rowCount, columnCount, 1, columnOffset, out successOne);
		List<string> bestRows = (successOne > successZero) ? oneBasedRows : zeroBasedRows;
		return string.Join(";", bestRows);
	}

	private static List<string> ReadFamilySizeTableRows(object table, int rowCount, int columnCount, int rowOffset, int columnOffset, out int successCount)
	{
		List<string> result = new List<string>();
		successCount = 0;
		if (table == null || rowCount <= 0 || columnCount <= 0)
		{
			return result;
		}
		for (int rowIndex = 0; rowIndex < rowCount; rowIndex++)
		{
			List<string> cells = new List<string>();
			for (int columnIndex = 0; columnIndex < columnCount; columnIndex++)
			{
				try
				{
					object value = InvokePublicInstanceMethod(table, "AsValueString", checked(rowIndex + rowOffset), checked(columnIndex + columnOffset));
					cells.Add(NormalizeLookupToken(Convert.ToString(value, CultureInfo.InvariantCulture)));
					successCount = checked(successCount + 1);
				}
				catch (Exception ex)
				{
					ProjectData.SetProjectError(ex);
					cells.Add("error:" + NormalizeLookupToken(ex.GetType().Name));
					ProjectData.ClearProjectError();
				}
			}
			result.Add("r" + rowIndex.ToString(CultureInfo.InvariantCulture) + "=" + string.Join(",", cells));
		}
		return result;
	}

	private static object InvokePublicInstanceMethod(object target, string methodName, params object[] args)
	{
		if (target == null || string.IsNullOrWhiteSpace(methodName))
		{
			return null;
		}
		MethodInfo method = target.GetType().GetMethods(BindingFlags.Public | BindingFlags.Instance).FirstOrDefault((MethodInfo x) => string.Equals(x.Name, methodName, StringComparison.Ordinal) && x.GetParameters().Length == (args?.Length ?? 0));
		return method?.Invoke(target, args ?? new object[0]);
	}

	private static string ReadStringProperty(object target, string propertyName)
	{
		try
		{
			object value = target?.GetType().GetProperty(propertyName, BindingFlags.Public | BindingFlags.Instance)?.GetValue(target, null);
			return Convert.ToString(value, CultureInfo.InvariantCulture) ?? string.Empty;
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ProjectData.ClearProjectError();
			return string.Empty;
		}
	}

	private static int ReadIntProperty(object target, string propertyName)
	{
		try
		{
			object value = target?.GetType().GetProperty(propertyName, BindingFlags.Public | BindingFlags.Instance)?.GetValue(target, null);
			return Convert.ToInt32(value, CultureInfo.InvariantCulture);
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ProjectData.ClearProjectError();
			return 0;
		}
	}

	private static string NormalizeLookupToken(string value)
	{
		return Normalize(value).Replace("\n", " ").Replace("\t", " ").Replace("|", "/").Replace(";", "/").Replace(",", "/").Trim();
	}

	private static string NormalizeLookupDisplayToken(string value)
	{
		return (value ?? string.Empty).Trim().Replace("\r\n", " ").Replace("\r", " ").Replace("\n", " ").Replace("\t", " ").Replace("|", "/").Replace(";", "/").Replace(",", "/").Trim();
	}

	private static void DisposeIfNeeded(object value)
	{
		try
		{
			if (value is IDisposable disposable)
			{
				disposable.Dispose();
			}
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ProjectData.ClearProjectError();
		}
	}

	private static string BuildFamilyParameterDefinitionSignature(FamilyParameter familyParameter)
	{
		if (familyParameter == null || familyParameter.Definition == null)
		{
			return string.Empty;
		}
		return Normalize(familyParameter.Definition.Name) + "|" + Normalize(familyParameter.StorageType.ToString()) + "|" + Normalize(SafeBool([SpecialName] () => familyParameter.IsInstance).ToString()) + "|" + Normalize(ResolveFamilyParameterFormula(familyParameter)) + "|" + Normalize(ResolveFamilyParameterGuid(familyParameter));
	}

	private static string BuildNestedFamilyLabelSignature(Document familyDocument, FamilyManager manager)
	{
		if (familyDocument == null || manager == null)
		{
			return string.Empty;
		}
		List<string> parts = new List<string>();
		try
		{
			List<FamilyInstance> instances = new FilteredElementCollector(familyDocument).OfClass(typeof(FamilyInstance)).Cast<FamilyInstance>().OrderBy([SpecialName] (FamilyInstance x) => Normalize(ResolveCategoryName(x)), StringComparer.Ordinal)
				.ThenBy([SpecialName] (FamilyInstance x) => Normalize(SafeElementName(x)), StringComparer.Ordinal)
				.ToList();
			foreach (FamilyInstance instance in instances)
			{
				string instanceLabel = BuildFamilyInstanceReferenceSignature(instance);
				foreach (Parameter parameter in (from Parameter x in instance.Parameters
					where ShouldCaptureParameter(x)
					select x).OrderBy([SpecialName] (Parameter x) => Normalize(ResolveParameterName(x)), StringComparer.Ordinal))
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
		return string.Join(";", parts.Distinct(StringComparer.Ordinal).OrderBy([SpecialName] (string x) => x, StringComparer.Ordinal));
	}

	private static string BuildNestedFamilyTypeParameterValueSignature(Document familyDocument, Parameter parameter)
	{
		string BuildNestedFamilyTypeParameterValueSignature;
		if (familyDocument == null || parameter == null)
		{
			BuildNestedFamilyTypeParameterValueSignature = string.Empty;
		}
		else
		{
			try
			{
				if (parameter.StorageType != StorageType.ElementId)
				{
					BuildNestedFamilyTypeParameterValueSignature = string.Empty;
				}
				else
				{
					ElementId valueId = parameter.AsElementId();
					if ((object)valueId == null || valueId == ElementId.InvalidElementId)
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
						else if (valueElement is FamilySymbol symbol)
						{
							string familyName = string.Empty;
							try
							{
								if (symbol.Family != null)
								{
									familyName = SafeElementName(symbol.Family);
								}
							}
							catch (Exception projectError)
							{
								ProjectData.SetProjectError(projectError);
								ProjectData.ClearProjectError();
							}
							BuildNestedFamilyTypeParameterValueSignature = "nested-family=" + Normalize(familyName) + "|nested-type=" + Normalize(SafeElementName(symbol)) + "|nested-category=" + Normalize(ResolveCategoryName(symbol));
						}
						else
						{
							BuildNestedFamilyTypeParameterValueSignature = "nested-family=|nested-type=" + Normalize(SafeElementName(valueElement)) + "|nested-category=" + Normalize(ResolveCategoryName(valueElement));
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
				symbolName = SafeElementName(instance.Symbol);
				if (instance.Symbol.Family != null)
				{
					familyName = SafeElementName(instance.Symbol.Family);
				}
			}
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ProjectData.ClearProjectError();
		}
		return Normalize(ResolveCategoryName(instance)) + "|" + Normalize(familyName) + "|" + Normalize(symbolName) + "|" + Normalize(SafeElementName(instance));
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
				MethodInfo methodInfo = manager.GetType().GetMethod("GetAssociatedFamilyParameter", new Type[1] { typeof(Parameter) });
				ResolveAssociatedFamilyParameter = (((object)methodInfo != null) ? (methodInfo.Invoke(manager, new object[1] { parameter }) as FamilyParameter) : null);
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
			ResolveFamilyParameterName = familyParameter?.Definition?.Name ?? string.Empty;
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
				PropertyInfo propertyInfo = familyParameter.GetType().GetProperty("Formula");
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
		if (idValue < 0 && !SafeBool([SpecialName] () => parameter.IsShared) && string.IsNullOrWhiteSpace(ResolveExternalGuid(parameter)))
		{
			return false;
		}
		return true;
	}

	private static int ResolveParameterIdInteger(Parameter parameter)
	{
		try
		{
			if (parameter != null && (object)parameter.Id != null)
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
		bool IsAnnotationFamily;
		try
		{
			IsAnnotationFamily = family != null && family.FamilyCategory != null && family.FamilyCategory.CategoryType == CategoryType.Annotation;
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
			ResolveParameterName = parameter?.Definition?.Name ?? string.Empty;
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
		string ResolveStorageTypeName;
		try
		{
			ResolveStorageTypeName = parameter.StorageType.ToString();
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
			if ((object)parameter.Id != null && RevitElementIdCompat.CompatIntegerValue(parameter.Id) < 0)
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
			if (parameter?.Definition is ExternalDefinition { GUID: var gUID })
			{
				return gUID.ToString("D");
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
			if (familyParameter.Definition is ExternalDefinition { GUID: var gUID })
			{
				return gUID.ToString("D");
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
			ResolveCategoryName = element.Category?.Name ?? string.Empty;
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
		try
		{
			Parameter parameterValue = ((Element)family).get_Parameter(BuiltInParameter.FAMILY_SHARED);
			if (parameterValue != null && parameterValue.StorageType == StorageType.Integer)
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
			if ((object)solid != null && solid.Edges != null)
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
