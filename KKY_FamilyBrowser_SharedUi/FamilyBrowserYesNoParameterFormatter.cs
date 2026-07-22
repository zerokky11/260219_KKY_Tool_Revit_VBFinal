using System;
using System.Globalization;
using System.Reflection;

public static class FamilyBrowserYesNoParameterFormatter
{
	public static string FormatInteger(object parameterOrDefinition, int value, string fallback = null)
	{
		if (IsYesNo(parameterOrDefinition))
		{
			return value == 0 ? "No" : "Yes";
		}

		return string.IsNullOrWhiteSpace(fallback)
			? value.ToString(CultureInfo.InvariantCulture)
			: fallback;
	}

	public static bool IsYesNo(object parameterOrDefinition)
	{
		if (parameterOrDefinition == null)
		{
			return false;
		}

		object definition = ReadProperty(parameterOrDefinition, "Definition") ?? parameterOrDefinition;
		string dataTypeId = ReadDataTypeId(definition);
		if (!string.IsNullOrWhiteSpace(dataTypeId) &&
			(dataTypeId.IndexOf("boolean", StringComparison.OrdinalIgnoreCase) >= 0 ||
			 dataTypeId.IndexOf("yesno", StringComparison.OrdinalIgnoreCase) >= 0))
		{
			return true;
		}

		object parameterType = ReadProperty(definition, "ParameterType");
		return parameterType != null &&
			string.Equals(parameterType.ToString(), "YesNo", StringComparison.OrdinalIgnoreCase);
	}

	private static string ReadDataTypeId(object definition)
	{
		try
		{
			MethodInfo getDataType = definition.GetType().GetMethod(
				"GetDataType",
				BindingFlags.Instance | BindingFlags.Public,
				null,
				Type.EmptyTypes,
				null);
			if (getDataType == null)
			{
				return string.Empty;
			}

			object dataType = getDataType.Invoke(definition, null);
			object typeId = ReadProperty(dataType, "TypeId");
			return typeId == null ? Convert.ToString(dataType, CultureInfo.InvariantCulture) : Convert.ToString(typeId, CultureInfo.InvariantCulture);
		}
		catch
		{
			return string.Empty;
		}
	}

	private static object ReadProperty(object source, string propertyName)
	{
		if (source == null)
		{
			return null;
		}

		try
		{
			PropertyInfo property = source.GetType().GetProperty(propertyName, BindingFlags.Instance | BindingFlags.Public);
			return property == null ? null : property.GetValue(source, null);
		}
		catch
		{
			return null;
		}
	}
}
