using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Security.Cryptography;
using System.Text;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.DB;
using Microsoft.VisualBasic.CompilerServices;

public sealed class SystemTypeSemanticCaptureService
{
	[CompilerGenerated]
	internal sealed class _Closure_0024__5_002D0
	{
		public string _0024VB_0024Local_systemFamilyKind;

		public string _0024VB_0024Local_typeName;

		public string _0024VB_0024Local_categoryName;

		public _Closure_0024__5_002D0(_Closure_0024__5_002D0 arg0)
		{
			if (arg0 != null)
			{
				_0024VB_0024Local_systemFamilyKind = arg0._0024VB_0024Local_systemFamilyKind;
				_0024VB_0024Local_typeName = arg0._0024VB_0024Local_typeName;
				_0024VB_0024Local_categoryName = arg0._0024VB_0024Local_categoryName;
			}
		}

		[SpecialName]
		internal bool _Lambda_0024__3(ElementType x)
		{
			return string.Equals(Normalize(((object)x).GetType().Name), Normalize(_0024VB_0024Local_systemFamilyKind), StringComparison.Ordinal);
		}

		[SpecialName]
		internal bool _Lambda_0024__4(ElementType x)
		{
			return string.Equals(Normalize(ResolveElementName((Element)(object)x)), Normalize(_0024VB_0024Local_typeName), StringComparison.Ordinal);
		}

		[SpecialName]
		internal bool _Lambda_0024__6(ElementType x)
		{
			return CategoryNamesMatch(ResolveCategoryName((Element)(object)x), _0024VB_0024Local_categoryName);
		}
	}

	private static readonly HashSet<string> AllowedSystemTypeNames = new HashSet<string>(new string[20]
	{
		"WallType", "FloorType", "RoofType", "CeilingType", "StairsType", "RailingType", "DuctType", "PipeType", "FlexDuctType", "FlexPipeType",
		"DuctSystemType", "PipingSystemType", "MechanicalSystemType", "ElectricalSystemType", "CableTrayType", "ConduitType", "WireType", "DuctInsulationType", "PipeInsulationType", "DuctLiningType"
	}, StringComparer.OrdinalIgnoreCase);

	private static readonly string[] ManagedSystemTypeParameterNameTokens = new string[12]
	{
		"classification", "uniformat", "assembly", "segment", "material", "shape", "profile", "system classification", "system abbreviation", "abbreviation",
		"service type", "service"
	};

	private SystemTypeSemanticCaptureService()
	{
	}

	public static SystemTypeCatalogSnapshot Capture(Document doc, string sourceId, IDictionary<string, string> loadableContentFingerprintCache = null, bool includeDeepLoadableContent = true, Action<int, int, string> progress = null)
	{
		//IL_0013: Unknown result type (might be due to invalid IL or missing references)
		SystemTypeCatalogSnapshot catalog = CreateCatalogShell(doc, sourceId);
		if (doc == null)
		{
			return catalog;
		}
		List<ElementType> systemTypes = (from ElementType x in (IEnumerable)new FilteredElementCollector(doc).WhereElementIsElementType()
			where x != null
			where !(x is FamilySymbol)
			where AllowedSystemTypeNames.Contains(((object)x).GetType().Name)
			select x).OrderBy<ElementType, string>([SpecialName] (ElementType x) => ((object)x).GetType().Name, StringComparer.Ordinal).ThenBy<ElementType, string>([SpecialName] (ElementType x) => Normalize(ResolveElementName((Element)(object)x)), StringComparer.Ordinal).ToList();
		int total = Math.Max(1, systemTypes.Count);
		checked
		{
			int num = systemTypes.Count - 1;
			for (int index = 0; index <= num; index++)
			{
				ElementType systemType = systemTypes[index];
				ReportProgress(progress, index + 1, total, "Reading system type " + (index + 1).ToString(CultureInfo.InvariantCulture) + "/" + total.ToString(CultureInfo.InvariantCulture) + ": " + ((object)systemType).GetType().Name + " / " + ResolveElementName((Element)(object)systemType));
				SystemTypeSemanticSnapshot snapshot = BuildSnapshot(doc, systemType, loadableContentFingerprintCache, includeDeepLoadableContent);
				catalog.Types.Add(snapshot);
			}
			return catalog;
		}
	}

	public static SystemTypeCatalogSnapshot CaptureSelected(Document doc, string sourceId, string systemFamilyKind, string categoryName, string typeName, IDictionary<string, string> loadableContentFingerprintCache = null, bool includeDeepLoadableContent = false, Action<int, int, string> progress = null)
	{
		//IL_004a: Unknown result type (might be due to invalid IL or missing references)
		_Closure_0024__5_002D0 arg = default(_Closure_0024__5_002D0);
		_Closure_0024__5_002D0 CS_0024_003C_003E8__locals8 = new _Closure_0024__5_002D0(arg);
		CS_0024_003C_003E8__locals8._0024VB_0024Local_systemFamilyKind = systemFamilyKind;
		CS_0024_003C_003E8__locals8._0024VB_0024Local_categoryName = categoryName;
		CS_0024_003C_003E8__locals8._0024VB_0024Local_typeName = typeName;
		SystemTypeCatalogSnapshot catalog = CreateCatalogShell(doc, sourceId);
		if (doc == null || string.IsNullOrWhiteSpace(CS_0024_003C_003E8__locals8._0024VB_0024Local_systemFamilyKind) || string.IsNullOrWhiteSpace(CS_0024_003C_003E8__locals8._0024VB_0024Local_typeName))
		{
			return catalog;
		}
		List<ElementType> candidates = (from ElementType x in (IEnumerable)new FilteredElementCollector(doc).WhereElementIsElementType()
			where x != null
			where !(x is FamilySymbol)
			where AllowedSystemTypeNames.Contains(((object)x).GetType().Name)
			where string.Equals(Normalize(((object)x).GetType().Name), Normalize(CS_0024_003C_003E8__locals8._0024VB_0024Local_systemFamilyKind), StringComparison.Ordinal)
			where string.Equals(Normalize(ResolveElementName((Element)(object)x)), Normalize(CS_0024_003C_003E8__locals8._0024VB_0024Local_typeName), StringComparison.Ordinal)
			select x).OrderBy<ElementType, string>([SpecialName] (ElementType x) => Normalize(ResolveCategoryName((Element)(object)x)), StringComparer.Ordinal).ToList();
		List<ElementType> categoryMatches = candidates.Where([SpecialName] (ElementType x) => CategoryNamesMatch(ResolveCategoryName((Element)(object)x), CS_0024_003C_003E8__locals8._0024VB_0024Local_categoryName)).ToList();
		if (categoryMatches.Count > 0)
		{
			candidates = categoryMatches;
		}
		int total = Math.Max(1, candidates.Count);
		checked
		{
			int num = candidates.Count - 1;
			for (int index = 0; index <= num; index++)
			{
				ElementType systemType = candidates[index];
				ReportProgress(progress, index + 1, total, "Reading selected system type " + (index + 1).ToString(CultureInfo.InvariantCulture) + "/" + total.ToString(CultureInfo.InvariantCulture) + ": " + ((object)systemType).GetType().Name + " / " + ResolveElementName((Element)(object)systemType));
				SystemTypeSemanticSnapshot snapshot = BuildSnapshot(doc, systemType, loadableContentFingerprintCache, includeDeepLoadableContent);
				catalog.Types.Add(snapshot);
			}
			return catalog;
		}
	}

	private static SystemTypeCatalogSnapshot CreateCatalogShell(Document doc, string sourceId)
	{
		SystemTypeCatalogSnapshot systemTypeCatalogSnapshot = new SystemTypeCatalogSnapshot();
		systemTypeCatalogSnapshot.SourceId = sourceId ?? string.Empty;
		systemTypeCatalogSnapshot.DocumentTitle = ((doc != null) ? doc.Title : null) ?? string.Empty;
		systemTypeCatalogSnapshot.DocumentPath = ((doc != null) ? doc.PathName : null) ?? string.Empty;
		systemTypeCatalogSnapshot.CapturedAtUtc = DateTime.UtcNow.ToString("O", CultureInfo.InvariantCulture);
		object obj;
		if (doc == null)
		{
			obj = null;
		}
		else
		{
			Application application = doc.Application;
			obj = ((application != null) ? application.VersionNumber : null);
		}
		if (obj == null)
		{
			obj = string.Empty;
		}
		systemTypeCatalogSnapshot.RevitVersion = (string)obj;
		return systemTypeCatalogSnapshot;
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

	private static SystemTypeSemanticSnapshot BuildSnapshot(Document doc, ElementType systemType, IDictionary<string, string> loadableContentFingerprintCache, bool includeDeepLoadableContent)
	{
		List<RoutingDependencySnapshot> routingDependencies = new List<RoutingDependencySnapshot>();
		Dictionary<string, string> parameters = CaptureRelevantParameters(systemType, doc);
		string routingPreferenceSignature = BuildRoutingPreferenceSignature(doc, systemType, routingDependencies, loadableContentFingerprintCache, includeDeepLoadableContent);
		string compoundStructureSignature = BuildCompoundStructureSignature(doc, systemType);
		SystemTypeSemanticSnapshot systemTypeSemanticSnapshot = new SystemTypeSemanticSnapshot();
		systemTypeSemanticSnapshot.SystemFamilyKind = ((object)systemType).GetType().Name;
		systemTypeSemanticSnapshot.CategoryName = ResolveCategoryName((Element)(object)systemType);
		systemTypeSemanticSnapshot.TypeName = ResolveElementName((Element)(object)systemType);
		systemTypeSemanticSnapshot.ClassificationCode = ResolveParameterCandidate(parameters, "classification", "uniformat", "assembly");
		systemTypeSemanticSnapshot.SegmentName = ResolveSegmentName(parameters, routingDependencies);
		systemTypeSemanticSnapshot.MaterialName = ResolveMaterialName(parameters, compoundStructureSignature);
		systemTypeSemanticSnapshot.Shape = ResolveShape(systemType, parameters);
		systemTypeSemanticSnapshot.RoutingPreferenceSignature = routingPreferenceSignature;
		systemTypeSemanticSnapshot.CompoundStructureSignature = compoundStructureSignature;
		systemTypeSemanticSnapshot.Parameters = parameters;
		systemTypeSemanticSnapshot.RoutingDependencies = routingDependencies;
		return systemTypeSemanticSnapshot;
	}

	private static Dictionary<string, string> CaptureRelevantParameters(ElementType systemType, Document doc)
	{
		Dictionary<string, string> result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
		foreach (Parameter parameter in ((IEnumerable)((Element)systemType).Parameters).Cast<Parameter>())
		{
			if (parameter != null && parameter.Definition != null && parameter.HasValue && ShouldCaptureSystemTypeParameter(parameter))
			{
				string key = parameter.Definition.Name;
				string value = ResolveParameterValue(parameter, doc);
				if (!string.IsNullOrWhiteSpace(key) && !string.IsNullOrWhiteSpace(value))
				{
					result[key.Trim()] = value.Trim();
				}
			}
		}
		return result;
	}

	private static bool ShouldCaptureSystemTypeParameter(Parameter parameter)
	{
		if (parameter == null || parameter.Definition == null)
		{
			return false;
		}
		if (ShouldSkipParameter(parameter))
		{
			return false;
		}
		if (IsSharedParameter(parameter))
		{
			return true;
		}
		if (HasManagedParameterToken(ResolveBuiltInParameterName(parameter)))
		{
			return true;
		}
		return HasManagedParameterToken(parameter.Definition.Name);
	}

	private static bool ShouldSkipParameter(Parameter parameter)
	{
		try
		{
			if (parameter.Id != null && RevitElementIdCompat.CompatIntegerValue(parameter.Id) == -1002001)
			{
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

	private static bool IsSharedParameter(Parameter parameter)
	{
		bool IsSharedParameter;
		try
		{
			if (parameter != null && parameter.IsShared)
			{
				IsSharedParameter = true;
				goto IL_0043;
			}
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ProjectData.ClearProjectError();
		}
		try
		{
			IsSharedParameter = ((parameter != null) ? parameter.Definition : null) is ExternalDefinition;
		}
		catch (Exception projectError2)
		{
			ProjectData.SetProjectError(projectError2);
			IsSharedParameter = false;
			ProjectData.ClearProjectError();
		}
		goto IL_0043;
		IL_0043:
		return IsSharedParameter;
	}

	private static string ResolveBuiltInParameterName(Parameter parameter)
	{
		//IL_002d: Unknown result type (might be due to invalid IL or missing references)
		string ResolveBuiltInParameterName;
		try
		{
			ResolveBuiltInParameterName = ((parameter != null && parameter.Id != null && RevitElementIdCompat.CompatIntegerValue(parameter.Id) < 0) ? ((Enum)(BuiltInParameter)RevitElementIdCompat.CompatIntegerValue(parameter.Id)/*cast due to .constrained prefix*/).ToString() : string.Empty);
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ResolveBuiltInParameterName = string.Empty;
			ProjectData.ClearProjectError();
		}
		return ResolveBuiltInParameterName;
	}

	private static bool HasManagedParameterToken(string value)
	{
		string normalized = Normalize(value).Replace("_", " ");
		if (string.IsNullOrWhiteSpace(normalized))
		{
			return false;
		}
		string[] managedSystemTypeParameterNameTokens = ManagedSystemTypeParameterNameTokens;
		foreach (string token in managedSystemTypeParameterNameTokens)
		{
			if (normalized.Contains(Normalize(token)))
			{
				return true;
			}
		}
		return false;
	}

	private static string ResolveParameterValue(Parameter parameter, Document doc)
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
				ResolveParameterValue = parameter.AsDouble().ToString("G17", CultureInfo.InvariantCulture);
				break;
			case 0:
				ResolveParameterValue = parameter.AsInteger().ToString(CultureInfo.InvariantCulture);
				break;
			case 3:
			{
				ElementId id = parameter.AsElementId();
				if (id == ElementId.InvalidElementId)
				{
					ResolveParameterValue = string.Empty;
					break;
				}
				if (RevitElementIdCompat.CompatIntegerValue(id) < 0)
				{
					ResolveParameterValue = RevitElementIdCompat.CompatIntegerValue(id).ToString(CultureInfo.InvariantCulture);
					break;
				}
				Element referenced = doc.GetElement(id);
				ResolveParameterValue = ((referenced != null) ? BuildReferencedElementParameterValue(doc, referenced) : RevitElementIdCompat.CompatIntegerValue(id).ToString(CultureInfo.InvariantCulture));
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

	private static string BuildReferencedElementParameterValue(Document doc, Element referenced)
	{
		if (referenced == null)
		{
			return string.Empty;
		}
		List<string> parts = new List<string>
		{
			((object)referenced).GetType().Name,
			ResolveCategoryName(referenced),
			ResolveElementName(referenced)
		};
		return NormalizeMultiline(string.Join("|", parts));
	}

	private unsafe static string BuildRoutingPreferenceSignature(Document doc, ElementType systemType, ICollection<RoutingDependencySnapshot> routingDependencies, IDictionary<string, string> loadableContentFingerprintCache, bool includeDeepLoadableContent)
	{
		//IL_00b9: Unknown result type (might be due to invalid IL or missing references)
		//IL_00be: Unknown result type (might be due to invalid IL or missing references)
		//IL_00c1: Unknown result type (might be due to invalid IL or missing references)
		//IL_010f: Unknown result type (might be due to invalid IL or missing references)
		//IL_021a: Unknown result type (might be due to invalid IL or missing references)
		//IL_0221: Expected O, but got Unknown
		PropertyInfo managerProperty = ((object)systemType).GetType().GetProperty("RoutingPreferenceManager");
		if ((object)managerProperty == null)
		{
			return string.Empty;
		}
		object? value = managerProperty.GetValue(systemType, null);
		RoutingPreferenceManager manager = (RoutingPreferenceManager)((value is RoutingPreferenceManager) ? value : null);
		if (manager == null)
		{
			return string.Empty;
		}
		List<string> lines = new List<string>();
		IOrderedEnumerable<RoutingPreferenceRuleGroupType> groups = from RoutingPreferenceRuleGroupType x in Enum.GetValues(typeof(RoutingPreferenceRuleGroupType))
			where (int)x != -1
			orderby (int)x
			select x;
		foreach (RoutingPreferenceRuleGroupType group in groups)
		{
			int ruleCount;
			try
			{
				ruleCount = manager.GetNumberOfRules(group);
			}
			catch (Exception projectError)
			{
				ProjectData.SetProjectError(projectError);
				lines.Add(Normalize(((Enum)(*(RoutingPreferenceRuleGroupType*)(&group))/*cast due to .constrained prefix*/).ToString()) + "|read-error|rule-count");
				ProjectData.ClearProjectError();
				continue;
			}
			int num = checked(ruleCount - 1);
			for (int index = 0; index <= num; index = checked(index + 1))
			{
				RoutingPreferenceRule rule = null;
				try
				{
					rule = manager.GetRule(group, index);
				}
				catch (Exception projectError2)
				{
					ProjectData.SetProjectError(projectError2);
					lines.Add(Normalize(((Enum)(*(RoutingPreferenceRuleGroupType*)(&group))/*cast due to .constrained prefix*/).ToString()) + "|" + index.ToString(CultureInfo.InvariantCulture) + "|read-error|rule");
					ProjectData.ClearProjectError();
					continue;
				}
				if (rule == null)
				{
					lines.Add(Normalize(((Enum)(*(RoutingPreferenceRuleGroupType*)(&group))/*cast due to .constrained prefix*/).ToString()) + "|" + index.ToString(CultureInfo.InvariantCulture) + "|null-rule");
					continue;
				}
				ElementId partId = rule.MEPPartId;
				if (RoutingRuleHasMappablePart(rule))
				{
					Element part = ((partId == null) ? null : doc.GetElement(partId));
					string partClass = ((part == null) ? string.Empty : ((object)part).GetType().Name);
					string partCategory = ResolveCategoryName(part);
					string familyKey = string.Empty;
					string familyName = string.Empty;
					string typeName = string.Empty;
					string familyFingerprint = string.Empty;
					string typeFingerprint = string.Empty;
					string partFingerprint = string.Empty;
					if (part is FamilySymbol)
					{
						FamilySymbol symbol = (FamilySymbol)part;
						Family family = symbol.Family;
						familyName = ((family != null) ? ((Element)family).Name : null) ?? string.Empty;
						typeName = ResolveElementName((Element)(object)symbol);
						familyKey = RoutingFamilyCatalogBuilder.BuildFamilyKey(ResolveFamilyCategoryName(family), familyName);
						familyFingerprint = BuildLoadableFamilyFingerprint(doc, family, loadableContentFingerprintCache, includeDeepLoadableContent);
						typeFingerprint = SystemTypeFingerprintService.ComputeSimpleTypeFingerprint(familyKey, typeName);
						routingDependencies.Add(new RoutingDependencySnapshot
						{
							DependencyRole = ((Enum)(*(RoutingPreferenceRuleGroupType*)(&group))/*cast due to .constrained prefix*/).ToString(),
							LibraryFamilyId = familyKey,
							FamilyName = familyName,
							TypeName = typeName,
							FamilyFingerprint = familyFingerprint,
							TypeFingerprint = typeFingerprint
						});
					}
					else
					{
						typeName = ResolveElementName(part);
						partFingerprint = BuildRoutingPartFingerprint(doc, part);
					}
					lines.Add(Normalize(((Enum)(*(RoutingPreferenceRuleGroupType*)(&group))/*cast due to .constrained prefix*/).ToString()) + "|" + index.ToString(CultureInfo.InvariantCulture) + "|" + Normalize(partClass) + "|" + Normalize(partCategory) + "|" + Normalize(familyKey) + "|" + Normalize(familyName) + "|" + Normalize(typeName) + "|" + Normalize(familyFingerprint) + "|" + Normalize(typeFingerprint) + "|" + Normalize(partFingerprint) + "|" + Normalize(BuildRoutingCriteriaSignature(rule)));
				}
			}
		}
		return string.Join("\n", lines);
	}

	private static string BuildRoutingPartFingerprint(Document doc, Element part)
	{
		string BuildRoutingPartFingerprint;
		if (part == null)
		{
			BuildRoutingPartFingerprint = string.Empty;
		}
		else
		{
			try
			{
				string signature = RoutingPartSignatureService.Build(doc, part);
				BuildRoutingPartFingerprint = ((!string.IsNullOrWhiteSpace(signature)) ? ("sha256:" + ComputeSha256(signature)) : string.Empty);
			}
			catch (Exception projectError)
			{
				ProjectData.SetProjectError(projectError);
				BuildRoutingPartFingerprint = string.Empty;
				ProjectData.ClearProjectError();
			}
		}
		return BuildRoutingPartFingerprint;
	}

	private static string ComputeSha256(string text)
	{
		byte[] bytes = Encoding.UTF8.GetBytes(text ?? string.Empty);
		using SHA256 sha = SHA256.Create();
		byte[] hash = sha.ComputeHash(bytes);
		StringBuilder sb = new StringBuilder(checked(hash.Length * 2));
		byte[] array = hash;
		foreach (byte b in array)
		{
			sb.Append(b.ToString("x2", CultureInfo.InvariantCulture));
		}
		return sb.ToString();
	}

	private static bool RoutingRuleHasMappablePart(RoutingPreferenceRule rule)
	{
		if (rule == null || rule.MEPPartId == null)
		{
			return false;
		}
		return rule.MEPPartId != ElementId.InvalidElementId && RevitElementIdCompat.CompatIntegerValue(rule.MEPPartId) > 0;
	}

	private static string BuildRoutingCriteriaSignature(RoutingPreferenceRule rule)
	{
		if (rule == null)
		{
			return string.Empty;
		}
		int count = TryGetCriterionCount(rule);
		if (count <= 0)
		{
			return string.Empty;
		}
		List<string> criteria = new List<string>();
		checked
		{
			int num = count - 1;
			for (int index = 0; index <= num; index++)
			{
				object criterion = RuntimeHelpers.GetObjectValue(TryGetCriterion(rule, index));
				if (criterion != null)
				{
					List<string> parts = new List<string> { criterion.GetType().Name };
					object minimumValue = RuntimeHelpers.GetObjectValue(TryGetPropertyValue(RuntimeHelpers.GetObjectValue(criterion), "MinimumSize"));
					if (minimumValue != null)
					{
						parts.Add("min=" + Convert.ToString(RuntimeHelpers.GetObjectValue(minimumValue), CultureInfo.InvariantCulture));
					}
					object maximumValue = RuntimeHelpers.GetObjectValue(TryGetPropertyValue(RuntimeHelpers.GetObjectValue(criterion), "MaximumSize"));
					if (maximumValue != null)
					{
						parts.Add("max=" + Convert.ToString(RuntimeHelpers.GetObjectValue(maximumValue), CultureInfo.InvariantCulture));
					}
					criteria.Add(string.Join(":", parts));
				}
			}
			return string.Join("&", criteria);
		}
	}

	private static int TryGetCriterionCount(RoutingPreferenceRule rule)
	{
		MethodInfo methodInfo = ((object)rule).GetType().GetMethod("GetNumberOfCriteria", Type.EmptyTypes);
		if ((object)methodInfo != null)
		{
			try
			{
				return Convert.ToInt32(RuntimeHelpers.GetObjectValue(methodInfo.Invoke(rule, null)), CultureInfo.InvariantCulture);
			}
			catch (Exception projectError)
			{
				ProjectData.SetProjectError(projectError);
				ProjectData.ClearProjectError();
			}
		}
		PropertyInfo propertyInfo = ((object)rule).GetType().GetProperty("NumberOfCriteria");
		if ((object)propertyInfo != null)
		{
			try
			{
				return Convert.ToInt32(RuntimeHelpers.GetObjectValue(propertyInfo.GetValue(rule, null)), CultureInfo.InvariantCulture);
			}
			catch (Exception projectError2)
			{
				ProjectData.SetProjectError(projectError2);
				ProjectData.ClearProjectError();
			}
		}
		return 0;
	}

	private static object TryGetCriterion(RoutingPreferenceRule rule, int index)
	{
		MethodInfo methodInfo = ((object)rule).GetType().GetMethod("GetCriterion", new Type[1] { typeof(int) });
		object TryGetCriterion;
		if ((object)methodInfo == null)
		{
			TryGetCriterion = null;
		}
		else
		{
			try
			{
				TryGetCriterion = methodInfo.Invoke(rule, new object[1] { index });
			}
			catch (Exception projectError)
			{
				ProjectData.SetProjectError(projectError);
				TryGetCriterion = null;
				ProjectData.ClearProjectError();
			}
		}
		return TryGetCriterion;
	}

	private static object TryGetPropertyValue(object instance, string propertyName)
	{
		object TryGetPropertyValue;
		if (instance == null)
		{
			TryGetPropertyValue = null;
		}
		else
		{
			PropertyInfo propertyInfo = instance.GetType().GetProperty(propertyName);
			if ((object)propertyInfo == null)
			{
				TryGetPropertyValue = null;
			}
			else
			{
				try
				{
					TryGetPropertyValue = propertyInfo.GetValue(RuntimeHelpers.GetObjectValue(instance), null);
				}
				catch (Exception projectError)
				{
					ProjectData.SetProjectError(projectError);
					TryGetPropertyValue = null;
					ProjectData.ClearProjectError();
				}
			}
		}
		return TryGetPropertyValue;
	}

	private static string BuildCompoundStructureSignature(Document doc, ElementType systemType)
	{
		//IL_00cc: Unknown result type (might be due to invalid IL or missing references)
		//IL_00d1: Unknown result type (might be due to invalid IL or missing references)
		HostObjAttributes hostAttributes = (HostObjAttributes)(object)((systemType is HostObjAttributes) ? systemType : null);
		if (hostAttributes == null)
		{
			return string.Empty;
		}
		CompoundStructure compound = hostAttributes.GetCompoundStructure();
		if (compound == null)
		{
			return string.Empty;
		}
		List<string> lines = new List<string> { "vertical=" + compound.IsVerticallyCompound.ToString().ToLowerInvariant() };
		checked
		{
			try
			{
				IList<CompoundStructureLayer> layers = compound.GetLayers();
				int num = layers.Count - 1;
				for (int index = 0; index <= num; index++)
				{
					CompoundStructureLayer layer = layers[index];
					string materialName = string.Empty;
					try
					{
						materialName = ResolveElementName(doc.GetElement(layer.MaterialId));
					}
					catch (Exception projectError)
					{
						ProjectData.SetProjectError(projectError);
						materialName = string.Empty;
						ProjectData.ClearProjectError();
					}
					lines.Add(index.ToString(CultureInfo.InvariantCulture) + "|" + Normalize(((Enum)layer.Function/*cast due to .constrained prefix*/).ToString()) + "|" + Normalize(materialName) + "|" + layer.Width.ToString("G17", CultureInfo.InvariantCulture));
				}
			}
			catch (Exception projectError2)
			{
				ProjectData.SetProjectError(projectError2);
				ProjectData.ClearProjectError();
			}
			return string.Join("\n", lines);
		}
	}

	private static string BuildLoadableFamilyFingerprint(Document doc, Family family, IDictionary<string, string> loadableContentFingerprintCache, bool includeDeepLoadableContent)
	{
		if (family == null)
		{
			return string.Empty;
		}
		List<string> typeNames = (from x in family.GetFamilySymbolIds().Select([SpecialName] (ElementId id) =>
			{
				Element element = doc.GetElement(id);
				return (ElementType)(object)((element is ElementType) ? element : null);
			})
			where x != null
			select ResolveElementName((Element)(object)x)).OrderBy<string, string>([SpecialName] (string x) => Normalize(x), StringComparer.Ordinal).ToList();
		return ProjectSnapshotFingerprintService.BuildLoadableFingerprint(new ProjectLoadableFamilySnapshotItem
		{
			FamilyName = (((Element)family).Name ?? string.Empty),
			CategoryName = ResolveFamilyCategoryName(family),
			CategoryGroup = ResolveFamilyCategoryGroup(family),
			TypeCount = typeNames.Count,
			TypeNames = typeNames,
			Parameters = CaptureLoadableFamilyParameters(doc, family),
			ContentFingerprint = LoadableFamilyContentSignatureService.Build(doc, family, loadableContentFingerprintCache, includeDeepLoadableContent),
			UniqueId = (((Element)family).UniqueId ?? string.Empty),
			IsShared = ResolveIsShared(family)
		});
	}

	private static List<StandardFamilyParameterSnapshotItem> CaptureLoadableFamilyParameters(Document doc, Family family)
	{
		List<StandardFamilyParameterSnapshotItem> result = new List<StandardFamilyParameterSnapshotItem>();
		if (doc == null || family == null)
		{
			return result;
		}
		foreach (Parameter parameter in (from Parameter x in (IEnumerable)((Element)family).Parameters
			where ShouldCaptureFamilyParameter(x)
			select x).OrderBy<Parameter, string>([SpecialName] (Parameter x) => Normalize(ResolveParameterName(x)), StringComparer.Ordinal))
		{
			result.Add(BuildFamilyParameterSnapshot(doc, parameter, "Family", string.Empty));
		}
		foreach (ElementId symbolId in family.GetFamilySymbolIds())
		{
			Element element = doc.GetElement(symbolId);
			FamilySymbol symbol = (FamilySymbol)(object)((element is FamilySymbol) ? element : null);
			if (symbol == null)
			{
				continue;
			}
			string typeName = ResolveElementName((Element)(object)symbol);
			foreach (Parameter parameter2 in (from Parameter x in (IEnumerable)((Element)symbol).Parameters
				where ShouldCaptureFamilyParameter(x)
				select x).OrderBy<Parameter, string>([SpecialName] (Parameter x) => Normalize(ResolveParameterName(x)), StringComparer.Ordinal))
			{
				result.Add(BuildFamilyParameterSnapshot(doc, parameter2, "Type", typeName));
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

	private static StandardFamilyParameterSnapshotItem BuildFamilyParameterSnapshot(Document doc, Parameter parameter, string scope, string typeName)
	{
		return new StandardFamilyParameterSnapshotItem
		{
			Scope = (scope ?? string.Empty),
			TypeName = (typeName ?? string.Empty),
			Name = ResolveParameterName(parameter),
			StorageType = ResolveStorageTypeName(parameter),
			ValuePreview = ResolveParameterValue(parameter, doc),
			Formula = string.Empty,
			IsInstance = false,
			IsReadOnly = SafeBool([SpecialName] () => parameter.IsReadOnly),
			IsShared = SafeBool([SpecialName] () => parameter.IsShared),
			ParameterId = ResolveParameterId(parameter),
			ExternalGuid = ResolveExternalGuid(parameter)
		};
	}

	private static bool ShouldCaptureFamilyParameter(Parameter parameter)
	{
		if (parameter == null || parameter.Definition == null)
		{
			return false;
		}
		if (string.IsNullOrWhiteSpace(ResolveParameterName(parameter)))
		{
			return false;
		}
		if (ShouldSkipParameter(parameter))
		{
			return false;
		}
		return parameter.HasValue;
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
		try
		{
			if (parameter != null && parameter.Id != null)
			{
				return RevitElementIdCompat.CompatIntegerValue(parameter.Id).ToString(CultureInfo.InvariantCulture);
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

	private static string ResolveParameterCandidate(IDictionary<string, string> parameters, params string[] keywords)
	{
		if (parameters == null)
		{
			return string.Empty;
		}
		foreach (KeyValuePair<string, string> kv in parameters)
		{
			string normalizedKey = Normalize(kv.Key);
			foreach (string keyword in keywords)
			{
				if (normalizedKey.Contains(Normalize(keyword)))
				{
					return kv.Value;
				}
			}
		}
		return string.Empty;
	}

	private static string ResolveSegmentName(IDictionary<string, string> parameters, IEnumerable<RoutingDependencySnapshot> routingDependencies)
	{
		string parameterValue = ResolveParameterCandidate(parameters, "segment", "세그먼트");
		if (!string.IsNullOrWhiteSpace(parameterValue))
		{
			return parameterValue;
		}
		RoutingDependencySnapshot segmentRule = routingDependencies.FirstOrDefault([SpecialName] (RoutingDependencySnapshot x) => string.Equals(Normalize(x.DependencyRole), "segments", StringComparison.Ordinal));
		if (segmentRule == null)
		{
			return string.Empty;
		}
		return segmentRule.TypeName;
	}

	private static string ResolveMaterialName(IDictionary<string, string> parameters, string compoundStructureSignature)
	{
		string parameterValue = ResolveParameterCandidate(parameters, "material", "재료");
		if (!string.IsNullOrWhiteSpace(parameterValue))
		{
			return parameterValue;
		}
		if (string.IsNullOrWhiteSpace(compoundStructureSignature))
		{
			return string.Empty;
		}
		string layerLine = compoundStructureSignature.Split(new string[1] { "\n" }, StringSplitOptions.RemoveEmptyEntries).FirstOrDefault([SpecialName] (string x) => x.Contains("|"));
		if (string.IsNullOrWhiteSpace(layerLine))
		{
			return string.Empty;
		}
		string[] tokens = layerLine.Split('|');
		if (tokens.Length < 2)
		{
			return string.Empty;
		}
		return tokens[1];
	}

	private static string ResolveShape(ElementType systemType, IDictionary<string, string> parameters)
	{
		PropertyInfo propertyInfo = ((object)systemType).GetType().GetProperty("Shape");
		if ((object)propertyInfo != null)
		{
			try
			{
				object value = RuntimeHelpers.GetObjectValue(propertyInfo.GetValue(systemType, null));
				if (value != null)
				{
					return Convert.ToString(RuntimeHelpers.GetObjectValue(value), CultureInfo.InvariantCulture);
				}
			}
			catch (Exception projectError)
			{
				ProjectData.SetProjectError(projectError);
				ProjectData.ClearProjectError();
			}
		}
		return ResolveParameterCandidate(parameters, "shape", "profile", "형상", "단면");
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

	private static string NormalizeMultiline(string value)
	{
		if (value == null)
		{
			return string.Empty;
		}
		return value.Trim().Replace("\r\n", "\n").Replace("\r", "\n");
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

	private static bool CategoryNamesMatch(string left, string right)
	{
		if (string.IsNullOrWhiteSpace(left) || string.IsNullOrWhiteSpace(right))
		{
			return string.IsNullOrWhiteSpace(left) && string.IsNullOrWhiteSpace(right);
		}
		return string.Equals(Normalize(left), Normalize(right), StringComparison.Ordinal);
	}
}
