using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Security.Cryptography;
using System.Text;
using Autodesk.Revit.DB;

public sealed class SystemTypeDetailedComponentSnapshotItem
{
	public int Order { get; set; }

	public string GroupName { get; set; }

	public string Path { get; set; }

	public string RoleName { get; set; }

	public string ComponentKind { get; set; }

	public string ReferenceClassName { get; set; }

	public string CategoryName { get; set; }

	public string FamilyName { get; set; }

	public string TypeName { get; set; }

	public string ValueKind { get; set; }

	public string RawValue { get; set; }

	public string DisplayValue { get; set; }

	public string ContentFingerprint { get; set; }

	public SystemTypeDetailedComponentSnapshotItem()
	{
		GroupName = string.Empty;
		Path = string.Empty;
		RoleName = string.Empty;
		ComponentKind = string.Empty;
		ReferenceClassName = string.Empty;
		CategoryName = string.Empty;
		FamilyName = string.Empty;
		TypeName = string.Empty;
		ValueKind = string.Empty;
		RawValue = string.Empty;
		DisplayValue = string.Empty;
		ContentFingerprint = string.Empty;
	}
}

public static class SystemTypeDetailedComponentSnapshotService
{
	private const string RequiredCurtainPanelGroupName = "CurtainPanelDependencies";

	private static readonly HashSet<string> SupportedRootTypes = new HashSet<string>(new[]
	{
		"RailingType",
		"StairsType"
	}, StringComparer.OrdinalIgnoreCase);

	private static readonly HashSet<string> RequiredCurtainRootTypes = new HashSet<string>(new[]
	{
		"WallType",
		"CurtainSystemType",
		"PanelType"
	}, StringComparer.OrdinalIgnoreCase);

	private static readonly BuiltInParameter[] CurtainPanelReferenceParameters = new[]
	{
		BuiltInParameter.AUTO_PANEL_WALL,
		BuiltInParameter.AUTO_PANEL
	};

	private static readonly HashSet<string> SupportedNestedObjects = new HashSet<string>(new[]
	{
		"BalusterPlacement",
		"BalusterPattern",
		"BalusterInfo",
		"PostPattern",
		"NonContinuousRailStructure",
		"NonContinuousRailInfo"
	}, StringComparer.OrdinalIgnoreCase);

	private static readonly HashSet<string> SkippedPropertyNames = new HashSet<string>(new[]
	{
		"Application",
		"BoundingBox",
		"Category",
		"DesignOption",
		"Document",
		"Geometry",
		"GroupId",
		"Id",
		"IsModifiable",
		"IsTransient",
		"IsValidObject",
		"Location",
		"Name",
		"OwnerViewId",
		"ParameterMap",
		"Parameters",
		"Pinned",
		"UniqueId",
		"ViewSpecific"
	}, StringComparer.OrdinalIgnoreCase);

	private static readonly string[] LengthTokens = new[]
	{
		"height", "width", "depth", "thickness", "offset", "spacing", "gap", "extension", "length", "distance", "nosing", "tread", "riser"
	};

	public static bool SupportsDetailedComponents(string typeClassName)
	{
		return SupportedRootTypes.Contains((typeClassName ?? string.Empty).Trim());
	}

	public static bool SupportsRequiredCurtainPanelComponents(string typeClassName)
	{
		return RequiredCurtainRootTypes.Contains((typeClassName ?? string.Empty).Trim());
	}

	public static bool IsRequiredCurtainPanelComponent(SystemTypeDetailedComponentSnapshotItem item)
	{
		return item != null && string.Equals((item.GroupName ?? string.Empty).Trim(), RequiredCurtainPanelGroupName, StringComparison.OrdinalIgnoreCase);
	}

	public static bool HasRequiredCurtainPanelComponents(IEnumerable<SystemTypeDetailedComponentSnapshotItem> items)
	{
		return (items ?? Enumerable.Empty<SystemTypeDetailedComponentSnapshotItem>()).Any(IsRequiredCurtainPanelComponent);
	}

	public static string BuildRequiredCurtainPanelSignature(IEnumerable<SystemTypeDetailedComponentSnapshotItem> items)
	{
		return BuildSignature((items ?? Enumerable.Empty<SystemTypeDetailedComponentSnapshotItem>()).Where(IsRequiredCurtainPanelComponent));
	}

	public static string BuildOptionalDetailedComponentSignature(IEnumerable<SystemTypeDetailedComponentSnapshotItem> items)
	{
		return BuildSignature((items ?? Enumerable.Empty<SystemTypeDetailedComponentSnapshotItem>()).Where(x => !IsRequiredCurtainPanelComponent(x)));
	}

	public static List<SystemTypeDetailedComponentSnapshotItem> Capture(
		Document document,
		ElementType systemType,
		IDictionary<string, string> loadableContentFingerprintCache,
		bool includeDeepLoadableContent,
		out bool captureCompleted)
	{
		captureCompleted = false;
		List<SystemTypeDetailedComponentSnapshotItem> result = new List<SystemTypeDetailedComponentSnapshotItem>();
		bool optionalRoot = document != null && systemType != null && SupportsDetailedComponents(systemType.GetType().Name);
		bool curtainRoot = document != null && systemType != null && IsCurtainDependencyRoot(systemType);
		if (!optionalRoot && !curtainRoot)
		{
			return result;
		}

		CaptureContext context = new CaptureContext(document, loadableContentFingerprintCache, includeDeepLoadableContent);
		try
		{
			if (optionalRoot)
			{
				CaptureObject(context, systemType, systemType.GetType().Name, systemType.GetType().Name, 0, result);
			}
			if (curtainRoot)
			{
				CaptureCurtainPanelDependencies(context, systemType, result);
			}
			result = result
				.Where(x => x != null && !string.IsNullOrWhiteSpace(x.Path))
				.GroupBy(x => BuildIdentityKey(x), StringComparer.Ordinal)
				.Select(x => x.First())
				.OrderBy(x => Normalize(x.GroupName), StringComparer.Ordinal)
				.ThenBy(x => Normalize(x.Path), StringComparer.Ordinal)
				.ThenBy(x => Normalize(x.RoleName), StringComparer.Ordinal)
				.ToList();
			for (int index = 0; index < result.Count; index++)
			{
				result[index].Order = index + 1;
			}
			// Referenced loadable-family content is only complete during the precise scan.
			// Fast scans may still expose rows for display, but must not certify a v4 comparison.
			captureCompleted = includeDeepLoadableContent;
		}
		catch
		{
			captureCompleted = false;
		}
		return result;
	}

	private static bool IsCurtainDependencyRoot(ElementType systemType)
	{
		if (systemType == null || !SupportsRequiredCurtainPanelComponents(systemType.GetType().Name))
		{
			return false;
		}
		if (string.Equals(systemType.GetType().Name, "CurtainSystemType", StringComparison.OrdinalIgnoreCase) ||
			string.Equals(systemType.GetType().Name, "PanelType", StringComparison.OrdinalIgnoreCase))
		{
			return true;
		}
		try
		{
			WallType wallType = systemType as WallType;
			return wallType != null && wallType.Kind == WallKind.Curtain;
		}
		catch
		{
			return false;
		}
	}

	private static void CaptureCurtainPanelDependencies(CaptureContext context, ElementType systemType, ICollection<SystemTypeDetailedComponentSnapshotItem> result)
	{
		string rootPath = systemType.GetType().Name + "/CurtainPanelConfiguration";
		result.Add(CreateValueItem(RequiredCurtainPanelGroupName, rootPath + "/CaptureStatus", "Curtain Panel Configuration", "Marker", "captured", "Captured"));
		FamilySymbol directPanelType = systemType as FamilySymbol;
		if (string.Equals(systemType.GetType().Name, "PanelType", StringComparison.OrdinalIgnoreCase) && directPanelType != null)
		{
			CaptureElement(context, directPanelType, RequiredCurtainPanelGroupName, rootPath + "/PanelType", "Curtain Panel Type", 0, result);
			return;
		}
		foreach (BuiltInParameter builtInParameter in CurtainPanelReferenceParameters)
		{
			Parameter parameter = null;
			try
			{
				parameter = systemType.get_Parameter(builtInParameter);
			}
			catch
			{
			}
			if (parameter == null)
			{
				continue;
			}
			string roleName = builtInParameter == BuiltInParameter.AUTO_PANEL_WALL ? "Default Curtain Panel" : "Curtain Panel";
			string path = rootPath + "/" + builtInParameter;
			CaptureCurtainParameter(context, parameter, path, roleName, 0, result);
		}
	}

	private static void CaptureCurtainParameter(CaptureContext context, Parameter parameter, string path, string roleName, int depth, ICollection<SystemTypeDetailedComponentSnapshotItem> result)
	{
		if (parameter == null || depth > 4)
		{
			return;
		}
		try
		{
			if (parameter.StorageType == StorageType.ElementId)
			{
				ElementId elementId = parameter.AsElementId();
				CaptureElementReference(context, elementId, RequiredCurtainPanelGroupName, path, roleName, depth, result);
				return;
			}
			string raw = ResolveParameterRawValue(parameter);
			string valueKind = ResolveParameterValueKind(parameter);
			string display = ResolveParameterDisplayValue(parameter, valueKind, raw);
			result.Add(CreateValueItem(RequiredCurtainPanelGroupName, path, roleName, valueKind, raw, display));
		}
		catch (Exception ex)
		{
			result.Add(CreateValueItem(RequiredCurtainPanelGroupName, path, roleName, "ReadError", ex.GetType().Name, ex.GetType().Name));
		}
	}

	private static void CaptureFamilySymbolTypeParameters(CaptureContext context, FamilySymbol symbol, string groupName, string path, int depth, ICollection<SystemTypeDetailedComponentSnapshotItem> result)
	{
		if (context == null || symbol == null || depth > 4)
		{
			return;
		}
		int symbolId = 0;
		try
		{
			symbolId = RevitElementIdCompat.CompatIntegerValue(symbol.Id);
		}
		catch
		{
		}
		if (symbolId > 0 && !context.VisitedElementIds.Add(symbolId))
		{
			return;
		}
		List<Parameter> parameters;
		try
		{
			parameters = symbol.Parameters.Cast<Parameter>()
				.Where(x => x != null && x.Definition != null && x.HasValue)
				.OrderBy(x => ResolveParameterIdentity(x), StringComparer.Ordinal)
				.ToList();
		}
		catch
		{
			return;
		}
		foreach (Parameter parameter in parameters)
		{
			string parameterIdentity = ResolveParameterIdentity(parameter);
			string parameterName = ResolveParameterName(parameter);
			string parameterPath = path + "/" + parameterIdentity;
			CaptureFamilySymbolParameter(context, parameter, groupName, parameterPath, parameterName, depth, result);
		}
	}

	private static void CaptureFamilySymbolParameter(CaptureContext context, Parameter parameter, string groupName, string path, string roleName, int depth, ICollection<SystemTypeDetailedComponentSnapshotItem> result)
	{
		if (parameter == null)
		{
			return;
		}
		try
		{
			if (parameter.StorageType == StorageType.ElementId)
			{
				CaptureElementReference(context, parameter.AsElementId(), groupName, path, roleName, depth, result);
				return;
			}
			string raw = ResolveParameterRawValue(parameter);
			string valueKind = ResolveParameterValueKind(parameter);
			string display = ResolveParameterDisplayValue(parameter, valueKind, raw);
			result.Add(CreateValueItem(groupName, path, roleName, valueKind, raw, display));
		}
		catch (Exception ex)
		{
			result.Add(CreateValueItem(groupName, path, roleName, "ReadError", ex.GetType().Name, ex.GetType().Name));
		}
	}

	private static string ResolveParameterIdentity(Parameter parameter)
	{
		if (parameter == null)
		{
			return "parameter";
		}
		try
		{
			int parameterId = RevitElementIdCompat.CompatIntegerValue(parameter.Id);
			if (parameterId < 0 && Enum.IsDefined(typeof(BuiltInParameter), parameterId))
			{
				return ((BuiltInParameter)parameterId).ToString();
			}
			if (parameter.IsShared)
			{
				return "shared-" + parameter.GUID.ToString("D", CultureInfo.InvariantCulture);
			}
		}
		catch
		{
		}
		return ResolveParameterName(parameter).Replace("/", "_");
	}

	private static string ResolveParameterName(Parameter parameter)
	{
		try
		{
			return parameter?.Definition?.Name ?? "Parameter";
		}
		catch
		{
			return "Parameter";
		}
	}

	private static string ResolveParameterRawValue(Parameter parameter)
	{
		switch (parameter.StorageType)
		{
			case StorageType.Double:
				return parameter.AsDouble().ToString("G17", CultureInfo.InvariantCulture);
			case StorageType.Integer:
				return parameter.AsInteger().ToString(CultureInfo.InvariantCulture);
			case StorageType.String:
				return parameter.AsString() ?? string.Empty;
			default:
				return string.Empty;
		}
	}

	private static string ResolveParameterValueKind(Parameter parameter)
	{
		if (parameter == null)
		{
			return "Text";
		}
		string spec = ResolveParameterSpecToken(parameter);
		if (parameter.StorageType == StorageType.Double)
		{
			if (spec.Contains("length"))
			{
				return "Length";
			}
			if (spec.Contains("angle"))
			{
				return "Angle";
			}
			return "Number";
		}
		if (parameter.StorageType == StorageType.Integer)
		{
			return spec.Contains("yesno") || spec.Contains("boolean") ? "Boolean" : "Integer";
		}
		if (parameter.StorageType == StorageType.String)
		{
			return "Text";
		}
		return "Value";
	}

	private static string ResolveParameterSpecToken(Parameter parameter)
	{
		List<string> tokens = new List<string>();
		try
		{
			object definition = parameter?.Definition;
			if (definition != null)
			{
				MethodInfo getDataType = definition.GetType().GetMethod("GetDataType", BindingFlags.Instance | BindingFlags.Public, null, Type.EmptyTypes, null);
				object dataType = getDataType?.Invoke(definition, null);
				if (dataType != null)
				{
					PropertyInfo typeId = dataType.GetType().GetProperty("TypeId", BindingFlags.Instance | BindingFlags.Public);
					tokens.Add(Convert.ToString(typeId?.GetValue(dataType, null), CultureInfo.InvariantCulture) ?? string.Empty);
					tokens.Add(Convert.ToString(dataType, CultureInfo.InvariantCulture) ?? string.Empty);
				}
				PropertyInfo parameterType = definition.GetType().GetProperty("ParameterType", BindingFlags.Instance | BindingFlags.Public);
				if (parameterType != null)
				{
					tokens.Add(Convert.ToString(parameterType.GetValue(definition, null), CultureInfo.InvariantCulture) ?? string.Empty);
				}
			}
		}
		catch
		{
		}
		return Normalize(string.Join(" ", tokens));
	}

	private static string ResolveParameterDisplayValue(Parameter parameter, string valueKind, string fallback)
	{
		try
		{
			if (string.Equals(valueKind, "Length", StringComparison.OrdinalIgnoreCase) && parameter.StorageType == StorageType.Double)
			{
				return (parameter.AsDouble() * 304.8).ToString("0.###", CultureInfo.InvariantCulture) + " mm";
			}
			if (string.Equals(valueKind, "Angle", StringComparison.OrdinalIgnoreCase) && parameter.StorageType == StorageType.Double)
			{
				return (parameter.AsDouble() * 180.0 / Math.PI).ToString("0.###", CultureInfo.InvariantCulture) + " deg";
			}
			if (string.Equals(valueKind, "Boolean", StringComparison.OrdinalIgnoreCase) && parameter.StorageType == StorageType.Integer)
			{
				return parameter.AsInteger() == 0 ? "No" : "Yes";
			}
		}
		catch
		{
		}
		try
		{
			string value = parameter.AsValueString();
			return string.IsNullOrWhiteSpace(value) ? fallback ?? string.Empty : value;
		}
		catch
		{
			return fallback ?? string.Empty;
		}
	}

	public static string BuildSignature(IEnumerable<SystemTypeDetailedComponentSnapshotItem> items)
	{
		List<string> lines = (items ?? Enumerable.Empty<SystemTypeDetailedComponentSnapshotItem>())
			.Where(x => x != null)
			.Select(x => string.Join("|", new[]
			{
				Normalize(x.GroupName),
				Normalize(x.Path),
				Normalize(x.RoleName),
				Normalize(x.ComponentKind),
				Normalize(x.ReferenceClassName),
				Normalize(x.CategoryName),
				Normalize(x.FamilyName),
				Normalize(x.TypeName),
				Normalize(x.ValueKind),
				NormalizeMultiline(x.RawValue),
				Normalize(x.ContentFingerprint)
			}))
			.OrderBy(x => x, StringComparer.Ordinal)
			.ToList();
		if (lines.Count == 0)
		{
			return string.Empty;
		}
		return "sha256:" + ComputeSha256("SystemTypeDetailedComponents|v1\n" + string.Join("\n", lines));
	}

	public static string BuildIdentityKey(SystemTypeDetailedComponentSnapshotItem item)
	{
		if (item == null)
		{
			return string.Empty;
		}
		return Normalize(item.GroupName) + "|" + Normalize(item.Path) + "|" + Normalize(item.RoleName);
	}

	public static string BuildComparableValue(SystemTypeDetailedComponentSnapshotItem item)
	{
		if (item == null)
		{
			return string.Empty;
		}
		return string.Join("|", new[]
		{
			Normalize(item.ComponentKind),
			Normalize(item.ReferenceClassName),
			Normalize(item.CategoryName),
			Normalize(item.FamilyName),
			Normalize(item.TypeName),
			Normalize(item.ValueKind),
			NormalizeMultiline(item.RawValue),
			Normalize(item.ContentFingerprint)
		});
	}

	public static string BuildDisplayValue(SystemTypeDetailedComponentSnapshotItem item)
	{
		if (item == null)
		{
			return string.Empty;
		}
		List<string> values = new List<string>();
		if (!string.IsNullOrWhiteSpace(item.FamilyName))
		{
			values.Add(item.FamilyName.Trim());
		}
		if (!string.IsNullOrWhiteSpace(item.TypeName) && !values.Any(x => string.Equals(x, item.TypeName.Trim(), StringComparison.OrdinalIgnoreCase)))
		{
			values.Add(item.TypeName.Trim());
		}
		if (values.Count == 0 && !string.IsNullOrWhiteSpace(item.DisplayValue))
		{
			values.Add(item.DisplayValue.Trim());
		}
		if (values.Count == 0 && !string.IsNullOrWhiteSpace(item.RawValue))
		{
			values.Add(item.RawValue.Trim());
		}
		return string.Join(" / ", values);
	}

	public static List<SystemTypeDetailedComponentSnapshotItem> CloneItems(IEnumerable<SystemTypeDetailedComponentSnapshotItem> items)
	{
		return (items ?? Enumerable.Empty<SystemTypeDetailedComponentSnapshotItem>())
			.Where(x => x != null)
			.Select(x => new SystemTypeDetailedComponentSnapshotItem
			{
				Order = x.Order,
				GroupName = x.GroupName ?? string.Empty,
				Path = x.Path ?? string.Empty,
				RoleName = x.RoleName ?? string.Empty,
				ComponentKind = x.ComponentKind ?? string.Empty,
				ReferenceClassName = x.ReferenceClassName ?? string.Empty,
				CategoryName = x.CategoryName ?? string.Empty,
				FamilyName = x.FamilyName ?? string.Empty,
				TypeName = x.TypeName ?? string.Empty,
				ValueKind = x.ValueKind ?? string.Empty,
				RawValue = x.RawValue ?? string.Empty,
				DisplayValue = x.DisplayValue ?? string.Empty,
				ContentFingerprint = x.ContentFingerprint ?? string.Empty
			})
			.ToList();
	}

	private static void CaptureObject(CaptureContext context, object instance, string groupName, string path, int depth, ICollection<SystemTypeDetailedComponentSnapshotItem> result)
	{
		if (context == null || instance == null || depth > 4 || result == null)
		{
			return;
		}
		Type instanceType = instance.GetType();
		if (!instanceType.IsValueType && !(instance is string) && !context.VisitedObjects.Add(instance))
		{
			return;
		}

		IEnumerable<PropertyInfo> properties = instanceType
			.GetProperties(BindingFlags.Instance | BindingFlags.Public)
			.Where(x => x != null && x.CanRead && x.GetIndexParameters().Length == 0)
			.Where(x => !SkippedPropertyNames.Contains(x.Name))
			.Where(x => x.DeclaringType != typeof(Element) && x.DeclaringType != typeof(ElementType))
			.OrderBy(x => x.Name, StringComparer.Ordinal);

		foreach (PropertyInfo property in properties)
		{
			string propertyPath = path + "/" + property.Name;
			object value;
			try
			{
				value = property.GetValue(instance, null);
			}
			catch (Exception ex)
			{
				result.Add(CreateValueItem(groupName, propertyPath, property.Name, "ReadError", ex.GetType().Name, ex.GetType().Name));
				continue;
			}
			CaptureValue(context, value, property.PropertyType, groupName, propertyPath, property.Name, depth, result);
		}

		CaptureIndexedCollection(context, instance, groupName, path, depth, result, "GetBalusterCount", "GetBaluster", "Baluster");
		CaptureIndexedCollection(context, instance, groupName, path, depth, result, "GetNonContinuousRailCount", "GetNonContinuousRail", "Rail");
	}

	private static void CaptureValue(CaptureContext context, object value, Type declaredType, string groupName, string path, string roleName, int depth, ICollection<SystemTypeDetailedComponentSnapshotItem> result)
	{
		if (value == null)
		{
			result.Add(CreateValueItem(groupName, path, roleName, "Empty", string.Empty, "-"));
			return;
		}
		Type valueType = value.GetType();
		if (value is ElementId)
		{
			CaptureElementReference(context, (ElementId)value, groupName, path, roleName, depth, result);
			return;
		}
		if (value is string || valueType.IsEnum || value is bool || IsNumericType(valueType))
		{
			string raw = FormatRawValue(value);
			string kind = ResolveValueKind(roleName, valueType);
			string display = FormatDisplayValue(roleName, value, kind, raw);
			result.Add(CreateValueItem(groupName, path, roleName, kind, raw, display));
			return;
		}
		if (SupportedNestedObjects.Contains(valueType.Name))
		{
			CaptureObject(context, value, groupName, path, depth + 1, result);
			return;
		}
		if (value is Element)
		{
			Element element = (Element)value;
			CaptureElement(context, element, groupName, path, roleName, depth, result);
			return;
		}
		if (declaredType != null && SupportedNestedObjects.Contains(declaredType.Name))
		{
			CaptureObject(context, value, groupName, path, depth + 1, result);
		}
	}

	private static void CaptureElementReference(CaptureContext context, ElementId elementId, string groupName, string path, string roleName, int depth, ICollection<SystemTypeDetailedComponentSnapshotItem> result)
	{
		int idValue = 0;
		try
		{
			idValue = RevitElementIdCompat.CompatIntegerValue(elementId);
		}
		catch
		{
		}
		if (elementId == null || elementId == ElementId.InvalidElementId || idValue <= 0)
		{
			result.Add(CreateValueItem(groupName, path, roleName, "ElementReference", idValue.ToString(CultureInfo.InvariantCulture), idValue == 0 ? "-" : idValue.ToString(CultureInfo.InvariantCulture)));
			return;
		}
		Element referenced = null;
		try
		{
			referenced = context.Document.GetElement(elementId);
		}
		catch
		{
		}
		if (referenced == null)
		{
			result.Add(CreateValueItem(groupName, path, roleName, "ElementReference", idValue.ToString(CultureInfo.InvariantCulture), "ElementId " + idValue.ToString(CultureInfo.InvariantCulture)));
			return;
		}
		CaptureElement(context, referenced, groupName, path, roleName, depth, result);
	}

	private static void CaptureElement(CaptureContext context, Element referenced, string groupName, string path, string roleName, int depth, ICollection<SystemTypeDetailedComponentSnapshotItem> result)
	{
		SystemTypeDetailedComponentSnapshotItem item = new SystemTypeDetailedComponentSnapshotItem
		{
			GroupName = groupName,
			Path = path,
			RoleName = roleName,
			ComponentKind = referenced is FamilySymbol ? "LoadableFamilyType" : "ElementReference",
			ReferenceClassName = referenced.GetType().Name,
			CategoryName = ResolveCategoryName(referenced),
			TypeName = ResolveElementName(referenced),
			ValueKind = "ElementReference",
			RawValue = ResolveElementName(referenced),
			DisplayValue = ResolveElementName(referenced)
		};
		if (referenced is FamilySymbol)
		{
			FamilySymbol symbol = (FamilySymbol)referenced;
			Family family = symbol.Family;
			item.FamilyName = family == null ? string.Empty : family.Name ?? string.Empty;
			item.TypeName = ResolveElementName(symbol);
			item.RawValue = item.FamilyName + "|" + item.TypeName;
			item.DisplayValue = BuildDisplayValue(item);
			try
			{
				item.ContentFingerprint = family == null
					? string.Empty
					: LoadableFamilyContentSignatureService.Build(context.Document, family, context.LoadableContentFingerprintCache, context.IncludeDeepLoadableContent);
			}
			catch
			{
				item.ContentFingerprint = string.Empty;
			}
		}
		result.Add(item);

		if (referenced is FamilySymbol)
		{
			if (depth < 3)
			{
				CaptureFamilySymbolTypeParameters(context, (FamilySymbol)referenced, groupName, path + "/TypeParameters", depth + 1, result);
			}
			return;
		}

		if (depth >= 3 || !(referenced is ElementType))
		{
			return;
		}
		int id = 0;
		try
		{
			id = RevitElementIdCompat.CompatIntegerValue(referenced.Id);
		}
		catch
		{
		}
		if (id > 0 && !context.VisitedElementIds.Add(id))
		{
			return;
		}
		CaptureObject(context, referenced, groupName, path + "/" + referenced.GetType().Name, depth + 1, result);
	}

	private static void CaptureIndexedCollection(CaptureContext context, object instance, string groupName, string path, int depth, ICollection<SystemTypeDetailedComponentSnapshotItem> result, string countMethodName, string itemMethodName, string itemLabel)
	{
		MethodInfo countMethod = instance.GetType().GetMethod(countMethodName, BindingFlags.Instance | BindingFlags.Public, null, Type.EmptyTypes, null);
		MethodInfo itemMethod = instance.GetType().GetMethod(itemMethodName, BindingFlags.Instance | BindingFlags.Public, null, new[] { typeof(int) }, null);
		if (countMethod == null || itemMethod == null)
		{
			return;
		}
		int count;
		try
		{
			count = Math.Max(0, Convert.ToInt32(countMethod.Invoke(instance, null), CultureInfo.InvariantCulture));
		}
		catch
		{
			return;
		}
		for (int index = 0; index < count; index++)
		{
			object value;
			try
			{
				value = itemMethod.Invoke(instance, new object[] { index });
			}
			catch (Exception ex)
			{
				result.Add(CreateValueItem(groupName, path + "/" + itemLabel + "[" + index.ToString(CultureInfo.InvariantCulture) + "]", itemLabel, "ReadError", ex.GetType().Name, ex.GetType().Name));
				continue;
			}
			CaptureObject(context, value, groupName, path + "/" + itemLabel + "[" + index.ToString(CultureInfo.InvariantCulture) + "]", depth + 1, result);
		}
	}

	private static SystemTypeDetailedComponentSnapshotItem CreateValueItem(string groupName, string path, string roleName, string valueKind, string rawValue, string displayValue)
	{
		return new SystemTypeDetailedComponentSnapshotItem
		{
			GroupName = groupName ?? string.Empty,
			Path = path ?? string.Empty,
			RoleName = roleName ?? string.Empty,
			ComponentKind = "Property",
			ValueKind = valueKind ?? string.Empty,
			RawValue = rawValue ?? string.Empty,
			DisplayValue = displayValue ?? string.Empty
		};
	}

	private static string ResolveValueKind(string propertyName, Type valueType)
	{
		if (valueType == typeof(bool))
		{
			return "Boolean";
		}
		if (valueType.IsEnum)
		{
			return "Enum";
		}
		if (valueType == typeof(double) || valueType == typeof(float) || valueType == typeof(decimal))
		{
			string normalized = Normalize(propertyName);
			if (LengthTokens.Any(x => normalized.Contains(x)))
			{
				return "Length";
			}
			if (normalized.Contains("angle") || normalized.Contains("slope"))
			{
				return "Angle";
			}
			return "Number";
		}
		if (IsNumericType(valueType))
		{
			return "Number";
		}
		return "Text";
	}

	private static string FormatRawValue(object value)
	{
		if (value == null)
		{
			return string.Empty;
		}
		if (value is bool)
		{
			return ((bool)value) ? "true" : "false";
		}
		if (value is double)
		{
			return ((double)value).ToString("G17", CultureInfo.InvariantCulture);
		}
		if (value is float)
		{
			return ((float)value).ToString("G9", CultureInfo.InvariantCulture);
		}
		if (value is IFormattable)
		{
			return ((IFormattable)value).ToString(null, CultureInfo.InvariantCulture);
		}
		return Convert.ToString(value, CultureInfo.InvariantCulture) ?? string.Empty;
	}

	private static string FormatDisplayValue(string propertyName, object value, string valueKind, string rawValue)
	{
		if (string.Equals(valueKind, "Length", StringComparison.Ordinal) && value is double)
		{
			return (((double)value) * 304.8).ToString("0.###", CultureInfo.InvariantCulture) + " mm";
		}
		if (string.Equals(valueKind, "Angle", StringComparison.Ordinal) && value is double)
		{
			return (((double)value) * 180.0 / Math.PI).ToString("0.###", CultureInfo.InvariantCulture) + " deg";
		}
		if (value is bool)
		{
			return ((bool)value) ? "Yes" : "No";
		}
		return rawValue ?? string.Empty;
	}

	private static bool IsNumericType(Type type)
	{
		TypeCode typeCode = Type.GetTypeCode(type);
		return typeCode == TypeCode.Byte || typeCode == TypeCode.SByte || typeCode == TypeCode.Int16 || typeCode == TypeCode.UInt16 || typeCode == TypeCode.Int32 || typeCode == TypeCode.UInt32 || typeCode == TypeCode.Int64 || typeCode == TypeCode.UInt64 || typeCode == TypeCode.Single || typeCode == TypeCode.Double || typeCode == TypeCode.Decimal;
	}

	private static string ResolveElementName(Element element)
	{
		try
		{
			return element == null ? string.Empty : element.Name ?? string.Empty;
		}
		catch
		{
			return string.Empty;
		}
	}

	private static string ResolveCategoryName(Element element)
	{
		try
		{
			return element == null || element.Category == null ? string.Empty : element.Category.Name ?? string.Empty;
		}
		catch
		{
			return string.Empty;
		}
	}

	private static string Normalize(string value)
	{
		return (value ?? string.Empty).Trim().ToLowerInvariant();
	}

	private static string NormalizeMultiline(string value)
	{
		return Normalize((value ?? string.Empty).Replace("\r\n", "\n").Replace("\r", "\n"));
	}

	private static string ComputeSha256(string value)
	{
		using (SHA256 sha = SHA256.Create())
		{
			byte[] hash = sha.ComputeHash(Encoding.UTF8.GetBytes(value ?? string.Empty));
			StringBuilder builder = new StringBuilder(hash.Length * 2);
			foreach (byte item in hash)
			{
				builder.Append(item.ToString("x2", CultureInfo.InvariantCulture));
			}
			return builder.ToString();
		}
	}

	private sealed class CaptureContext
	{
		public Document Document { get; private set; }

		public IDictionary<string, string> LoadableContentFingerprintCache { get; private set; }

		public bool IncludeDeepLoadableContent { get; private set; }

		public HashSet<int> VisitedElementIds { get; private set; }

		public HashSet<object> VisitedObjects { get; private set; }

		public CaptureContext(Document document, IDictionary<string, string> loadableContentFingerprintCache, bool includeDeepLoadableContent)
		{
			Document = document;
			LoadableContentFingerprintCache = loadableContentFingerprintCache ?? new Dictionary<string, string>(StringComparer.Ordinal);
			IncludeDeepLoadableContent = includeDeepLoadableContent;
			VisitedElementIds = new HashSet<int>();
			VisitedObjects = new HashSet<object>(ReferenceEqualityComparer.Instance);
		}
	}

	private sealed class ReferenceEqualityComparer : IEqualityComparer<object>
	{
		public static readonly ReferenceEqualityComparer Instance = new ReferenceEqualityComparer();

		public new bool Equals(object x, object y)
		{
			return ReferenceEquals(x, y);
		}

		public int GetHashCode(object obj)
		{
			return obj == null ? 0 : RuntimeHelpers.GetHashCode(obj);
		}
	}
}
