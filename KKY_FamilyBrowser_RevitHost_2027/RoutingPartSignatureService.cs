using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using Autodesk.Revit.DB;
using Microsoft.VisualBasic.CompilerServices;

public sealed class RoutingPartSignatureService
{
	private RoutingPartSignatureService()
	{
	}

	public static string Build(Document doc, Element element)
	{
		if (element == null)
		{
			return string.Empty;
		}
		return BuildCore(doc, element, 0, new HashSet<int>());
	}

	private static string BuildCore(Document doc, Element element, int depth, ISet<int> visitedElementIds)
	{
		if (element == null)
		{
			return string.Empty;
		}
		int elementId = ResolveElementIntegerId(element);
		if (elementId >= 0 && visitedElementIds != null && visitedElementIds.Contains(elementId))
		{
			return "class=" + Normalize(((object)element).GetType().Name) + "\ncategory=" + Normalize(ResolveCategoryName(element)) + "\nname=" + Normalize(ResolveElementName(element)) + "\ncycle=true";
		}
		if (elementId >= 0)
		{
			visitedElementIds?.Add(elementId);
		}
		List<string> lines = new List<string>
		{
			"class=" + Normalize(((object)element).GetType().Name),
			"category=" + Normalize(ResolveCategoryName(element)),
			"name=" + Normalize(ResolveElementName(element)),
			"params=" + BuildParameterSignature(doc, element, depth, visitedElementIds),
			"sizes=" + BuildSizeTableSignature(element)
		};
		if (elementId >= 0)
		{
			visitedElementIds?.Remove(elementId);
		}
		return string.Join("\n", lines);
	}

	public static bool Matches(Document sourceDocument, Element sourceElement, Document targetDocument, Element targetElement)
	{
		return string.Equals(Normalize(Build(sourceDocument, sourceElement)), Normalize(Build(targetDocument, targetElement)), StringComparison.Ordinal);
	}

	private static string BuildParameterSignature(Document doc, Element element, int depth, ISet<int> visitedElementIds)
	{
		if (element == null)
		{
			return string.Empty;
		}
		List<string> parts = new List<string>();
		try
		{
			foreach (Parameter parameter in (from Parameter x in (IEnumerable)element.Parameters
				where ShouldCaptureParameter(x)
				select x).OrderBy<Parameter, string>([SpecialName] (Parameter x) => Normalize(ResolveParameterName(x)), StringComparer.Ordinal))
			{
				parts.Add(Normalize(ResolveParameterName(parameter)) + ":" + Normalize(ResolveStorageTypeName(parameter)) + ":" + Normalize(ResolveParameterValue(doc, parameter, depth, visitedElementIds)) + ":" + Normalize(BuildPortableParameterIdentity(parameter)));
			}
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ProjectData.ClearProjectError();
		}
		return string.Join("|", parts);
	}

	private static string BuildSizeTableSignature(Element element)
	{
		if (element == null)
		{
			return string.Empty;
		}
		List<string> sizeSignatures = new List<string>();
		IOrderedEnumerable<MethodInfo> methods = (from x in ((object)element).GetType().GetMethods(BindingFlags.Instance | BindingFlags.Public)
			where x.GetParameters().Length == 0
			where string.Equals(x.Name, "GetSizes", StringComparison.OrdinalIgnoreCase) || string.Equals(x.Name, "GetMEPSizes", StringComparison.OrdinalIgnoreCase)
			select x).OrderBy<MethodInfo, string>([SpecialName] (MethodInfo x) => x.Name, StringComparer.Ordinal);
		foreach (MethodInfo methodInfo in methods)
		{
			try
			{
				if (!(RuntimeHelpers.GetObjectValue(methodInfo.Invoke(element, null)) is IEnumerable values))
				{
					continue;
				}
				foreach (object item in values)
				{
					string signature = BuildObjectSignature(RuntimeHelpers.GetObjectValue(RuntimeHelpers.GetObjectValue(item)));
					if (!string.IsNullOrWhiteSpace(signature))
					{
						sizeSignatures.Add(Normalize(methodInfo.Name) + ":" + signature);
					}
				}
			}
			catch (Exception projectError)
			{
				ProjectData.SetProjectError(projectError);
				ProjectData.ClearProjectError();
			}
		}
		return string.Join("|", sizeSignatures.OrderBy<string, string>([SpecialName] (string x) => x, StringComparer.Ordinal));
	}

	private static string BuildObjectSignature(object value)
	{
		if (value == null)
		{
			return string.Empty;
		}
		List<string> parts = new List<string>();
		IOrderedEnumerable<PropertyInfo> properties = (from x in value.GetType().GetProperties(BindingFlags.Instance | BindingFlags.Public)
			where x.GetIndexParameters().Length == 0
			select x).OrderBy<PropertyInfo, string>([SpecialName] (PropertyInfo x) => x.Name, StringComparer.Ordinal);
		foreach (PropertyInfo propertyInfo in properties)
		{
			try
			{
				string formatted = FormatSignatureValue(RuntimeHelpers.GetObjectValue(RuntimeHelpers.GetObjectValue(propertyInfo.GetValue(RuntimeHelpers.GetObjectValue(value), null))));
				if (!string.IsNullOrWhiteSpace(formatted))
				{
					parts.Add(Normalize(propertyInfo.Name) + "=" + Normalize(formatted));
				}
			}
			catch (Exception projectError)
			{
				ProjectData.SetProjectError(projectError);
				ProjectData.ClearProjectError();
			}
		}
		return string.Join(";", parts);
	}

	private static string FormatSignatureValue(object value)
	{
		//IL_0017: Unknown result type (might be due to invalid IL or missing references)
		//IL_0021: Expected O, but got Unknown
		if (value == null)
		{
			return string.Empty;
		}
		if (value is ElementId)
		{
			return RevitElementIdCompat.CompatIntegerValue((ElementId)value).ToString(CultureInfo.InvariantCulture);
		}
		if (value is double)
		{
			return Conversions.ToDouble(value).ToString("G17", CultureInfo.InvariantCulture);
		}
		if (value is float)
		{
			return Conversions.ToSingle(value).ToString("G9", CultureInfo.InvariantCulture);
		}
		if (value is decimal)
		{
			return Conversions.ToDecimal(value).ToString(CultureInfo.InvariantCulture);
		}
		if (value is IFormattable)
		{
			return ((IFormattable)value).ToString(null, CultureInfo.InvariantCulture);
		}
		return Convert.ToString(RuntimeHelpers.GetObjectValue(value), CultureInfo.InvariantCulture);
	}

	private static bool ShouldCaptureParameter(Parameter parameter)
	{
		if (parameter == null || parameter.Definition == null)
		{
			return false;
		}
		try
		{
			if (parameter.Id != null && RevitElementIdCompat.CompatIntegerValue(parameter.Id) == -1002001)
			{
				return false;
			}
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ProjectData.ClearProjectError();
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

	private static string ResolveParameterValue(Document doc, Parameter parameter, int depth, ISet<int> visitedElementIds)
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
					Element referenced = ((doc == null) ? null : doc.GetElement(id));
					ResolveParameterValue = ((referenced != null) ? BuildReferencedElementValue(doc, referenced, depth, visitedElementIds) : RevitElementIdCompat.CompatIntegerValue(id).ToString(CultureInfo.InvariantCulture));
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

	private static string BuildReferencedElementValue(Document doc, Element referenced, int depth, ISet<int> visitedElementIds)
	{
		if (referenced == null)
		{
			return string.Empty;
		}
		string header = ((object)referenced).GetType().Name + ":" + ResolveCategoryName(referenced) + ":" + ResolveElementName(referenced);
		if (depth >= 1)
		{
			return header;
		}
		return header + "{" + BuildCore(doc, referenced, checked(depth + 1), visitedElementIds) + "}";
	}

	private static int ResolveElementIntegerId(Element element)
	{
		try
		{
			if (element != null && element.Id != null)
			{
				return RevitElementIdCompat.CompatIntegerValue(element.Id);
			}
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ProjectData.ClearProjectError();
		}
		return -1;
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
}
