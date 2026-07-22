using System;
using System.Collections.Generic;
using System.Globalization;
using System.Reflection;
using Autodesk.Revit.DB;
using Microsoft.VisualBasic.CompilerServices;

public sealed class FamilyBrowserFamilyClassificationService
{
	public const string CategoryGroupModel = "Model";

	public const string CategoryGroupAnnotation = "Annotation";

	public const string CategoryGroupOther = "Other";

	private static readonly HashSet<int> AnnotationCategoryIds = BuildCategoryIdSet("OST_GenericAnnotation", "OST_GridHeads", "OST_LevelHeads", "OST_SectionHeads", "OST_ElevationMarks", "OST_CalloutHeads", "OST_Viewers", "OST_ViewReference", "OST_DoorTags", "OST_WindowTags", "OST_WallTags", "OST_RoomTags", "OST_SpaceTags", "OST_AreaTags", "OST_KeynoteTags", "OST_MultiCategoryTags", "OST_PipeTags", "OST_DuctTags", "OST_CableTrayTags", "OST_ConduitTags");

	private static readonly HashSet<int> RoutingDependencyCategoryIds = BuildCategoryIdSet("OST_DuctFitting", "OST_DuctAccessory", "OST_PipeFitting", "OST_PipeAccessory", "OST_CableTrayFitting", "OST_ConduitFitting");

	private static readonly HashSet<int> TypeManagedCategoryIds = BuildCategoryIdSet("OST_CurtainWallPanels", "OST_CurtainWallMullions", "OST_Mullions", "OST_WallSweep", "OST_Reveals");

	private FamilyBrowserFamilyClassificationService()
	{
	}

	public static bool IsBrowserLoadableFamily(Family family)
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
		if (!IsEditableFamily(family))
		{
			return false;
		}
		if (IsTypeManagedFamilyCategory(family))
		{
			return false;
		}
		return true;
	}

	public static string ResolveCategoryName(Family family)
	{
		string ResolveCategoryName;
		try
		{
			ResolveCategoryName = ((family == null || family.FamilyCategory == null) ? (family?.Category?.Name ?? string.Empty) : (family.FamilyCategory.Name ?? string.Empty));
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ResolveCategoryName = string.Empty;
			ProjectData.ClearProjectError();
		}
		return ResolveCategoryName;
	}

	public static string ResolveCategoryId(Family family)
	{
		try
		{
			if (family != null && family.FamilyCategory != null && (object)family.FamilyCategory.Id != null)
			{
				return RevitElementIdCompat.CompatIntegerValue(family.FamilyCategory.Id).ToString(CultureInfo.InvariantCulture);
			}
			if (family != null && family.Category != null && (object)family.Category.Id != null)
			{
				return RevitElementIdCompat.CompatIntegerValue(family.Category.Id).ToString(CultureInfo.InvariantCulture);
			}
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ProjectData.ClearProjectError();
		}
		return string.Empty;
	}

	public static string ResolveCategoryGroup(Family family)
	{
		string ResolveCategoryGroup;
		if (family == null)
		{
			ResolveCategoryGroup = string.Empty;
		}
		else
		{
			try
			{
				if (family.FamilyCategory == null)
				{
					ResolveCategoryGroup = "Other";
				}
				else
				{
					string categoryName = ResolveCategoryName(family);
					string categoryId = ResolveCategoryId(family);
					string familyName = family.Name ?? string.Empty;
					ResolveCategoryGroup = (IsAnnotationCategoryLike(categoryName, categoryId, familyName) ? "Annotation" : (family.FamilyCategory.CategoryType switch
					{
						CategoryType.Model => "Model",
						CategoryType.Annotation => "Annotation",
						_ => "Other",
					}));
				}
			}
			catch (Exception projectError)
			{
				ProjectData.SetProjectError(projectError);
				ResolveCategoryGroup = "Other";
				ProjectData.ClearProjectError();
			}
		}
		return ResolveCategoryGroup;
	}

	public static string ResolveCategoryGroup(string categoryGroup, string categoryName, string categoryId, string familyName)
	{
		if (IsAnnotationCategoryLike(categoryName, categoryId, familyName))
		{
			return "Annotation";
		}
		string normalizedGroup = NormalizeCategoryGroup(categoryGroup);
		if (normalizedGroup.Length > 0)
		{
			return normalizedGroup;
		}
		if (!string.IsNullOrWhiteSpace(categoryName))
		{
			return "Model";
		}
		return "Other";
	}

	public static bool IsRoutingDependencyFamilyCategory(Family family)
	{
		int? categoryId = ResolveCategoryIntegerId(family);
		if (categoryId.HasValue && RoutingDependencyCategoryIds.Contains(categoryId.Value))
		{
			return true;
		}
		return IsRoutingDependencyCategoryName(ResolveCategoryName(family));
	}

	public static bool IsRoutingDependencyCategory(string categoryName, string categoryId)
	{
		if (int.TryParse((categoryId ?? string.Empty).Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out var parsedCategoryId) && RoutingDependencyCategoryIds.Contains(parsedCategoryId))
		{
			return true;
		}
		return IsRoutingDependencyCategoryName(categoryName);
	}

	public static bool IsTypeManagedFamilyLike(string categoryName, string categoryId, string familyName)
	{
		if (int.TryParse((categoryId ?? string.Empty).Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out var parsedCategoryId) && TypeManagedCategoryIds.Contains(parsedCategoryId))
		{
			return true;
		}
		return IsTypeManagedFamilyName(categoryName, familyName);
	}

	private static bool IsRoutingDependencyCategoryName(string categoryName)
	{
		string compact = Normalize(categoryName).Replace(" ", string.Empty);
		if (compact.Length == 0)
		{
			return false;
		}
		return compact.Contains("ductfitting") || compact.Contains("ductaccessory") || compact.Contains("pipefitting") || compact.Contains("pipeaccessory") || compact.Contains("cabletrayfitting") || compact.Contains("conduitfitting");
	}

	private static bool IsAnnotationCategoryLike(string categoryName, string categoryId, string familyName)
	{
		if (int.TryParse((categoryId ?? string.Empty).Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out var parsedCategoryId) && AnnotationCategoryIds.Contains(parsedCategoryId))
		{
			return true;
		}
		string categoryCompact = Compact(categoryName);
		string familyCompact = Compact(familyName);
		if (categoryCompact.Contains("annotation") || categoryCompact.Contains("annotationsymbol") || categoryCompact.Contains("genericannotation") || categoryCompact.Contains("tag") || categoryCompact.Contains("text") || categoryCompact.Contains("viewreference") || categoryCompact.Contains("revision") || categoryCompact.Contains("titleblock") || categoryCompact.Contains("주석") || categoryCompact.Contains("태그") || categoryCompact.Contains("기호") || categoryCompact.Contains("문자") || categoryCompact.Contains("뷰참조"))
		{
			return true;
		}
		return familyCompact.Contains("gridhead") || familyCompact.Contains("levelhead") || familyCompact.Contains("sectionhead") || familyCompact.Contains("callouthead") || familyCompact.Contains("elevationmark") || familyCompact.Contains("그리드헤드") || familyCompact.Contains("레벨헤드");
	}

	private static bool IsEditableFamily(Family family)
	{
		try
		{
			PropertyInfo prop = family.GetType().GetProperty("IsEditable", BindingFlags.Instance | BindingFlags.Public);
			if ((object)prop != null && (object)prop.PropertyType == typeof(bool))
			{
				return Conversions.ToBoolean(prop.GetValue(family, null));
			}
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ProjectData.ClearProjectError();
		}
		return true;
	}

	private static bool IsTypeManagedFamilyCategory(Family family)
	{
		return IsTypeManagedFamilyLike(ResolveCategoryName(family), ResolveCategoryId(family), family?.Name ?? string.Empty);
	}

	private static bool IsTypeManagedFamilyName(string categoryName, string familyName)
	{
		string categoryCompact = Compact(categoryName);
		string familyCompact = Compact(familyName);
		string combinedCompact = categoryCompact + familyCompact;
		if (combinedCompact.Length == 0)
		{
			return false;
		}
		if (combinedCompact.Contains("mullion"))
		{
			return true;
		}
		if (combinedCompact.Contains("멀리언") || combinedCompact.Contains("몰리언"))
		{
			return true;
		}
		if (combinedCompact.Contains("systempanel"))
		{
			return true;
		}
		if (combinedCompact.Contains("시스템패널"))
		{
			return true;
		}
		if (combinedCompact.Contains("curtainwallpanel") || combinedCompact.Contains("curtainpanel"))
		{
			return true;
		}
		if (combinedCompact.Contains("커튼월패널") || combinedCompact.Contains("커튼패널"))
		{
			return true;
		}
		return categoryCompact.Contains("wallsweep") || categoryCompact.Contains("wallreveal") || categoryCompact.Contains("reveal") || categoryCompact.Contains("벽스윕") || categoryCompact.Contains("벽리빌") || categoryCompact.Contains("벽드러냄");
	}

	private static int? ResolveCategoryIntegerId(Family family)
	{
		try
		{
			if (family != null && family.FamilyCategory != null && (object)family.FamilyCategory.Id != null)
			{
				return RevitElementIdCompat.CompatIntegerValue(family.FamilyCategory.Id);
			}
			if (family != null && family.Category != null && (object)family.Category.Id != null)
			{
				return RevitElementIdCompat.CompatIntegerValue(family.Category.Id);
			}
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ProjectData.ClearProjectError();
		}
		return null;
	}

	private static HashSet<int> BuildCategoryIdSet(params string[] names)
	{
		HashSet<int> result = new HashSet<int>();
		foreach (string name in names)
		{
			try
			{
				BuiltInCategory value = (BuiltInCategory)Enum.Parse(typeof(BuiltInCategory), name, ignoreCase: false);
				result.Add((int)value);
			}
			catch (Exception projectError)
			{
				ProjectData.SetProjectError(projectError);
				ProjectData.ClearProjectError();
			}
		}
		return result;
	}

	private static string Normalize(string value)
	{
		return (value ?? string.Empty).Trim().ToLowerInvariant();
	}

	private static string Compact(string value)
	{
		return Normalize(value).Replace(" ", string.Empty).Replace("_", string.Empty).Replace("-", string.Empty);
	}

	private static string NormalizeCategoryGroup(string value)
	{
		switch (Normalize(value))
		{
		case "model":
		case "모델":
			return "Model";
		case "annotation":
		case "annotationcategory":
		case "주석":
		case "주석카테고리":
			return "Annotation";
		case "other":
		case "기타":
			return "Other";
		default:
			return string.Empty;
		}
	}
}
