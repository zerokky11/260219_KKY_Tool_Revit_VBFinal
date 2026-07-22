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
			return "class=" + Normalize(element.GetType().Name) + "\ncategory=" + Normalize(ResolveCategoryName(element)) + "\nname=" + Normalize(ResolveElementName(element)) + "\ncycle=true";
		}
		if (elementId >= 0)
		{
			visitedElementIds?.Add(elementId);
		}
		List<string> lines = new List<string>
		{
			"class=" + Normalize(element.GetType().Name),
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
			foreach (Parameter parameter in (from Parameter x in element.Parameters
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
		IOrderedEnumerable<MethodInfo> methods = (from x in element.GetType().GetMethods(BindingFlags.Instance | BindingFlags.Public)
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
			if ((object)parameter.Id != null && RevitElementIdCompat.CompatIntegerValue(parameter.Id) == -1002001)
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

	private static string ResolveParameterValue(Document doc, Parameter parameter, int depth, ISet<int> visitedElementIds)
	{
		string ResolveParameterValue;
		try
		{
			switch (parameter.StorageType)
			{
			case StorageType.String:
				ResolveParameterValue = NormalizeMultiline(parameter.AsString());
				break;
			case StorageType.Double:
			{
				string formatted2 = parameter.AsValueString();
				ResolveParameterValue = (string.IsNullOrWhiteSpace(formatted2) ? parameter.AsDouble().ToString("G17", CultureInfo.InvariantCulture) : NormalizeMultiline(formatted2));
				break;
			}
			case StorageType.Integer:
			{
				string formatted3 = parameter.AsValueString();
				ResolveParameterValue = (string.IsNullOrWhiteSpace(formatted3) ? parameter.AsInteger().ToString(CultureInfo.InvariantCulture) : NormalizeMultiline(formatted3));
				break;
			}
			case StorageType.ElementId:
			{
				ElementId id = parameter.AsElementId();
				if ((object)id == null || id == ElementId.InvalidElementId)
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
					Element referenced = doc?.GetElement(id);
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
		string header = referenced.GetType().Name + ":" + ResolveCategoryName(referenced) + ":" + ResolveElementName(referenced);
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
			if (element != null && (object)element.Id != null)
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
