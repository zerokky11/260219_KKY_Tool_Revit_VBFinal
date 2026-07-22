using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using Autodesk.Revit.DB;

public sealed class SystemTypeDetailSummaryService
{
	private SystemTypeDetailSummaryService()
	{
	}

	public static string Build(Document doc, ElementType systemType, SystemTypeSemanticSnapshot snapshot, IEnumerable<StandardSystemTypeLayerSnapshotItem> layers = null)
	{
		List<string> lines = new List<string> { "@system-detail-v1" };
		AddSection(lines, "identity");
		AddRow(lines, "identity", "Class", systemType?.GetType().Name ?? snapshot?.SystemFamilyKind);
		AddRow(lines, "identity", "Category", ResolveCategoryName(systemType) ?? snapshot?.CategoryName);
		AddRow(lines, "identity", "Type", ResolveElementName(systemType) ?? snapshot?.TypeName);
		AddRow(lines, "identity", "Classification", snapshot?.ClassificationCode);
		AddRow(lines, "identity", "Shape", snapshot?.Shape);
		AddRoutingRuleRows(lines, doc, systemType);
		AddDependencyRows(lines, snapshot?.RoutingDependencies);
		AddLayerRows(lines, layers);
		AddDetailedComponentRows(lines, snapshot);
		return string.Join(Environment.NewLine, lines.Where(x => !string.IsNullOrWhiteSpace(x)));
	}

	private static void AddDetailedComponentRows(List<string> lines, SystemTypeSemanticSnapshot snapshot)
	{
		if (snapshot == null || !snapshot.DetailedComponentsCaptured)
		{
			return;
		}
		List<SystemTypeDetailedComponentSnapshotItem> allRows = (snapshot.DetailedComponents ?? new List<SystemTypeDetailedComponentSnapshotItem>())
			.Where(x => x != null)
			.OrderBy(x => x.Order)
			.ThenBy(x => Normalize(x.Path), StringComparer.Ordinal)
			.ToList();
		if (SystemTypeDetailedComponentSnapshotService.SupportsDetailedComponents(snapshot.SystemFamilyKind))
		{
			AddDetailedComponentSection(lines, "components", allRows.Where(x => !SystemTypeDetailedComponentSnapshotService.IsRequiredCurtainPanelComponent(x)), "Configuration");
		}
		if (SystemTypeDetailedComponentSnapshotService.HasRequiredCurtainPanelComponents(allRows))
		{
			AddDetailedComponentSection(lines, "curtain-components", allRows.Where(SystemTypeDetailedComponentSnapshotService.IsRequiredCurtainPanelComponent), "Default Panel");
		}
	}

	private static void AddDetailedComponentSection(List<string> lines, string section, IEnumerable<SystemTypeDetailedComponentSnapshotItem> sourceRows, string emptyLabel)
	{
		AddSection(lines, section);
		List<SystemTypeDetailedComponentSnapshotItem> rows = (sourceRows ?? Enumerable.Empty<SystemTypeDetailedComponentSnapshotItem>())
			.Where(x => x != null && !string.Equals(x.ValueKind, "Marker", StringComparison.OrdinalIgnoreCase))
			.ToList();
		if (rows.Count == 0)
		{
			AddRow(lines, section, emptyLabel, "None");
			return;
		}
		foreach (SystemTypeDetailedComponentSnapshotItem row in rows)
		{
			string reference = JoinClean(" / ", row.ReferenceClassName, row.CategoryName);
			string value = JoinClean(" | ", SystemTypeDetailedComponentSnapshotService.BuildDisplayValue(row), reference, row.Path);
			string name = string.IsNullOrWhiteSpace(row.RoleName) ? row.ComponentKind : row.RoleName;
			AddStructuredDetailedComponentRow(lines, section, name, row, reference);
			AddRow(lines, section, name, value);
		}
	}

	private static void AddRoutingRuleRows(List<string> lines, Document doc, ElementType systemType)
	{
		List<string[]> routingRows = new List<string[]>();
		List<string[]> routingPreferenceRows = new List<string[]>();
		List<string[]> segmentRows = new List<string[]>();
		RoutingPreferenceManager manager = ResolveRoutingPreferenceManager(systemType);
		if (manager != null && doc != null)
		{
			foreach (RoutingPreferenceRuleGroupType group in Enum.GetValues(typeof(RoutingPreferenceRuleGroupType)).Cast<RoutingPreferenceRuleGroupType>().Where(x => x != RoutingPreferenceRuleGroupType.Undefined).OrderBy(x => (int)x))
			{
				int ruleCount = SafeGetRuleCount(manager, group);
				for (int index = 0; index < ruleCount; index++)
				{
					RoutingPreferenceRule rule = SafeGetRule(manager, group, index);
					if (rule == null || rule.MEPPartId == null || rule.MEPPartId == ElementId.InvalidElementId || RevitElementIdCompat.CompatIntegerValue(rule.MEPPartId) <= 0)
					{
						continue;
					}
					Element part = doc.GetElement(rule.MEPPartId);
					string groupName = group.ToString();
					string partClass = part?.GetType().Name ?? string.Empty;
					string partCategory = ResolveCategoryName(part);
					string partName = ResolvePartDisplayName(part);
					int sizeCount = CountElementSizes(part);
					string criteria = BuildRoutingCriteriaDisplay(rule);
					string value = BuildRoutingRuleValue(partClass, partCategory, partName, sizeCount, criteria);
					routingRows.Add(new[] { "routing", groupName + "[" + index.ToString(CultureInfo.InvariantCulture) + "]", value });
					routingPreferenceRows.Add(new[]
					{
						groupName,
						index.ToString(CultureInfo.InvariantCulture),
						partName,
						partClass,
						partCategory,
						sizeCount >= 0 ? sizeCount.ToString(CultureInfo.InvariantCulture) : "-",
						criteria
					});
					if (IsSegmentRule(groupName, partClass, partCategory))
					{
						segmentRows.Add(new[] { "segments", string.IsNullOrWhiteSpace(partName) ? groupName : partName, BuildSegmentValue(partClass, partCategory, sizeCount, criteria) });
					}
				}
			}
		}
		AddSection(lines, "routing");
		foreach (string[] route in routingPreferenceRows)
		{
			AddRoutingPreferenceRow(lines, route[0], route[1], route[2], route[3], route[4], route[5], route[6]);
		}
		if (routingRows.Count == 0)
		{
			AddRow(lines, "routing", "Rules", "None");
		}
		else
		{
			foreach (string[] row in routingRows)
			{
				AddRow(lines, row[0], row[1], row[2]);
			}
		}
		AddSection(lines, "segments");
		if (segmentRows.Count == 0)
		{
			AddRow(lines, "segments", "Segments", "None");
		}
		else
		{
			foreach (string[] row in segmentRows.GroupBy(x => Normalize(x[1]) + "|" + Normalize(x[2]), StringComparer.Ordinal).Select(x => x.First()).OrderBy(x => Normalize(x[1]), StringComparer.Ordinal))
			{
				AddRow(lines, row[0], row[1], row[2]);
			}
		}
	}

	private static void AddDependencyRows(List<string> lines, IEnumerable<RoutingDependencySnapshot> dependencies)
	{
		AddSection(lines, "dependencies");
		List<RoutingDependencySnapshot> rows = (dependencies ?? Enumerable.Empty<RoutingDependencySnapshot>()).Where(x => x != null && (!string.IsNullOrWhiteSpace(x.FamilyName) || !string.IsNullOrWhiteSpace(x.TypeName))).GroupBy(x => Normalize(x.DependencyRole) + "|" + Normalize(x.FamilyName) + "|" + Normalize(x.TypeName), StringComparer.Ordinal).Select(x => x.First()).OrderBy(x => Normalize(x.DependencyRole), StringComparer.Ordinal).ThenBy(x => Normalize(x.FamilyName), StringComparer.Ordinal).ThenBy(x => Normalize(x.TypeName), StringComparer.Ordinal).ToList();
		if (rows.Count == 0)
		{
			AddRow(lines, "dependencies", "Families", "None");
			return;
		}
		Dictionary<string, int> routeIndexByRole = new Dictionary<string, int>(StringComparer.Ordinal);
		foreach (RoutingDependencySnapshot dependency in rows)
		{
			string value = dependency.FamilyName;
			if (!string.IsNullOrWhiteSpace(dependency.TypeName))
			{
				value = string.IsNullOrWhiteSpace(value) ? dependency.TypeName : value + " / " + dependency.TypeName;
			}
			string roleKey = Normalize(dependency.DependencyRole);
			int routeIndex;
			if (!routeIndexByRole.TryGetValue(roleKey, out routeIndex))
			{
				routeIndex = 0;
			}
			routeIndexByRole[roleKey] = routeIndex + 1;
			AddRoutingPreferenceRow(lines, dependency.DependencyRole, routeIndex.ToString(CultureInfo.InvariantCulture), value, "FamilySymbol", string.Empty, "-", string.Empty);
			AddRow(lines, "dependencies", dependency.DependencyRole, value);
		}
	}

	private static void AddLayerRows(List<string> lines, IEnumerable<StandardSystemTypeLayerSnapshotItem> layers)
	{
		List<StandardSystemTypeLayerSnapshotItem> rows = (layers ?? Enumerable.Empty<StandardSystemTypeLayerSnapshotItem>()).Where(x => x != null).OrderBy(x => x.Index <= 0 ? int.MaxValue : x.Index).ToList();
		if (rows.Count == 0)
		{
			return;
		}
		AddSection(lines, "layers");
		int fallbackIndex = 1;
		foreach (StandardSystemTypeLayerSnapshotItem layer in rows)
		{
			int displayIndex = layer.Index > 0 ? layer.Index : fallbackIndex;
			AddStructuredLayerRow(lines, displayIndex, layer);
			AddRow(lines, "layers", "#" + displayIndex.ToString(CultureInfo.InvariantCulture), JoinClean(" / ", layer.FunctionName, layer.MaterialName, layer.ThicknessDisplay));
			fallbackIndex++;
		}
	}

	private static RoutingPreferenceManager ResolveRoutingPreferenceManager(ElementType systemType)
	{
		try
		{
			return systemType?.GetType().GetProperty("RoutingPreferenceManager")?.GetValue(systemType, null) as RoutingPreferenceManager;
		}
		catch
		{
			return null;
		}
	}

	private static int SafeGetRuleCount(RoutingPreferenceManager manager, RoutingPreferenceRuleGroupType group)
	{
		try
		{
			return Math.Max(0, manager.GetNumberOfRules(group));
		}
		catch
		{
			return 0;
		}
	}

	private static RoutingPreferenceRule SafeGetRule(RoutingPreferenceManager manager, RoutingPreferenceRuleGroupType group, int index)
	{
		try
		{
			return manager.GetRule(group, index);
		}
		catch
		{
			return null;
		}
	}

	private static string BuildRoutingRuleValue(string partClass, string partCategory, string partName, int sizeCount, string criteria)
	{
		List<string> parts = new List<string>();
		if (!string.IsNullOrWhiteSpace(partName))
		{
			parts.Add(partName);
		}
		parts.Add(JoinClean(" / ", partClass, partCategory));
		if (sizeCount >= 0)
		{
			parts.Add("sizes=" + sizeCount.ToString(CultureInfo.InvariantCulture));
		}
		if (!string.IsNullOrWhiteSpace(criteria))
		{
			parts.Add("criteria=" + criteria);
		}
		return JoinClean(" | ", parts.ToArray());
	}

	private static string BuildSegmentValue(string partClass, string partCategory, int sizeCount, string criteria)
	{
		List<string> parts = new List<string> { JoinClean(" / ", partClass, partCategory) };
		if (sizeCount >= 0)
		{
			parts.Add("sizes=" + sizeCount.ToString(CultureInfo.InvariantCulture));
		}
		if (!string.IsNullOrWhiteSpace(criteria))
		{
			parts.Add("criteria=" + criteria);
		}
		return JoinClean(" | ", parts.ToArray());
	}

	private static string BuildRoutingCriteriaDisplay(RoutingPreferenceRule rule)
	{
		int count = TryGetCriterionCount(rule);
		if (count <= 0)
		{
			return string.Empty;
		}
		List<string> criteria = new List<string>();
		for (int index = 0; index < count; index++)
		{
			object criterion = TryGetCriterion(rule, index);
			if (criterion == null)
			{
				continue;
			}
			string min = FormatCriterionValue(TryGetPropertyValue(criterion, "MinimumSize"));
			string max = FormatCriterionValue(TryGetPropertyValue(criterion, "MaximumSize"));
			criteria.Add(JoinClean(" ",
				"size " + (index + 1).ToString(CultureInfo.InvariantCulture),
				string.IsNullOrWhiteSpace(min) ? string.Empty : "min=" + min,
				string.IsNullOrWhiteSpace(max) ? string.Empty : "max=" + max));
		}
		return string.Join("; ", criteria);
	}

	private static int TryGetCriterionCount(RoutingPreferenceRule rule)
	{
		if (rule == null)
		{
			return 0;
		}
		MethodInfo method = rule.GetType().GetMethod("GetNumberOfCriteria", Type.EmptyTypes);
		if (method != null)
		{
			try
			{
				return Convert.ToInt32(method.Invoke(rule, null), CultureInfo.InvariantCulture);
			}
			catch
			{
			}
		}
		PropertyInfo property = rule.GetType().GetProperty("NumberOfCriteria");
		if (property != null)
		{
			try
			{
				return Convert.ToInt32(property.GetValue(rule, null), CultureInfo.InvariantCulture);
			}
			catch
			{
			}
		}
		return 0;
	}

	private static object TryGetCriterion(RoutingPreferenceRule rule, int index)
	{
		MethodInfo method = rule?.GetType().GetMethod("GetCriterion", new[] { typeof(int) });
		if (method == null)
		{
			return null;
		}
		try
		{
			return method.Invoke(rule, new object[] { index });
		}
		catch
		{
			return null;
		}
	}

	private static object TryGetPropertyValue(object instance, string propertyName)
	{
		PropertyInfo property = instance?.GetType().GetProperty(propertyName);
		if (property == null)
		{
			return null;
		}
		try
		{
			return property.GetValue(instance, null);
		}
		catch
		{
			return null;
		}
	}

	private static string FormatCriterionValue(object value)
	{
		if (value == null)
		{
			return string.Empty;
		}
		if (value is IFormattable formattable)
		{
			return formattable.ToString(null, CultureInfo.InvariantCulture);
		}
		return Convert.ToString(RuntimeHelpers.GetObjectValue(value), CultureInfo.InvariantCulture);
	}

	private static int CountElementSizes(Element element)
	{
		if (element == null)
		{
			return -1;
		}
		foreach (MethodInfo method in element.GetType().GetMethods(BindingFlags.Instance | BindingFlags.Public).Where(x => x.GetParameters().Length == 0 && (string.Equals(x.Name, "GetSizes", StringComparison.OrdinalIgnoreCase) || string.Equals(x.Name, "GetMEPSizes", StringComparison.OrdinalIgnoreCase))).OrderBy(x => x.Name, StringComparer.Ordinal))
		{
			try
			{
				if (method.Invoke(element, null) is IEnumerable values)
				{
					int count = 0;
					foreach (object value in values)
					{
						if (value != null)
						{
							count++;
						}
					}
					return count;
				}
			}
			catch
			{
			}
		}
		return -1;
	}

	private static bool IsSegmentRule(string groupName, string partClass, string partCategory)
	{
		string value = Normalize(groupName + " " + partClass + " " + partCategory);
		return value.Contains("segment");
	}

	private static string ResolvePartDisplayName(Element part)
	{
		if (part is FamilySymbol symbol)
		{
			string familyName = symbol.Family?.Name ?? string.Empty;
			string typeName = ResolveElementName(symbol);
			return JoinClean(" / ", familyName, typeName);
		}
		return ResolveElementName(part);
	}

	private static string ResolveElementName(Element element)
	{
		if (element == null)
		{
			return string.Empty;
		}
		try
		{
			return element.Name ?? string.Empty;
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
			return element?.Category?.Name ?? string.Empty;
		}
		catch
		{
			return string.Empty;
		}
	}

	private static void AddSection(List<string> lines, string section)
	{
		lines.Add("@section\t" + Clean(section));
	}

	private static void AddRoutingPreferenceRow(List<string> lines, string group, string priority, string partName, string partClass, string partCategory, string sizeCount, string criteria)
	{
		lines.Add(string.Join("\t", new[]
		{
			"@route",
			Clean(group),
			Clean(priority),
			Clean(partName),
			Clean(partClass),
			Clean(partCategory),
			Clean(string.IsNullOrWhiteSpace(sizeCount) ? "-" : sizeCount),
			Clean(criteria)
		}));
	}

	private static void AddStructuredLayerRow(List<string> lines, int displayIndex, StandardSystemTypeLayerSnapshotItem layer)
	{
		if (layer == null)
		{
			return;
		}
		lines.Add(string.Join("\t", new[]
		{
			"@layer",
			displayIndex.ToString(CultureInfo.InvariantCulture),
			Clean(layer.FunctionName),
			Clean(layer.MaterialName),
			layer.ThicknessFeet.ToString("G17", CultureInfo.InvariantCulture),
			Clean(layer.ThicknessDisplay),
			layer.IsCore ? "true" : "false",
			layer.IsStructuralMaterial ? "true" : "false",
			layer.IsVariable ? "true" : "false"
		}));
	}

	private static void AddRow(List<string> lines, string section, string name, string value)
	{
		if (string.IsNullOrWhiteSpace(name) && string.IsNullOrWhiteSpace(value))
		{
			return;
		}
		lines.Add("@row\t" + Clean(section) + "\t" + Clean(name) + "\t" + Clean(string.IsNullOrWhiteSpace(value) ? "-" : value));
	}

	private static void AddStructuredDetailedComponentRow(List<string> lines, string section, string name, SystemTypeDetailedComponentSnapshotItem row, string reference)
	{
		if (row == null)
		{
			return;
		}
		lines.Add(string.Join("\t", new[]
		{
			"@component",
			Clean(section),
			Clean(name),
			Clean(row.ValueKind),
			Clean(row.RawValue),
			Clean(SystemTypeDetailedComponentSnapshotService.BuildDisplayValue(row)),
			Clean(reference),
			Clean(row.Path)
		}));
	}

	private static string JoinClean(string separator, params string[] values)
	{
		return string.Join(separator, (values ?? Array.Empty<string>()).Where(x => !string.IsNullOrWhiteSpace(x)).Select(x => x.Trim()));
	}

	private static string Clean(string value)
	{
		return (value ?? string.Empty).Replace("\r\n", " ").Replace("\r", " ").Replace("\n", " ").Replace("\t", " ").Trim();
	}

	private static string Normalize(string value)
	{
		return (value ?? string.Empty).Trim().ToLowerInvariant();
	}
}
