using System;
using System.Collections;
using System.Globalization;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Text;
using Microsoft.VisualBasic.CompilerServices;

public sealed class PlainJsonReportWriter
{
	private PlainJsonReportWriter()
	{
	}

	public static string Serialize(object value)
	{
		StringBuilder stringBuilder = new StringBuilder(4096);
		WriteValue(stringBuilder, RuntimeHelpers.GetObjectValue(value));
		return stringBuilder.ToString();
	}

	private static void WriteValue(StringBuilder builder, object value)
	{
		if (value == null)
		{
			builder.Append("null");
			return;
		}
		Type valueType = value.GetType();
		if (value is string)
		{
			WriteString(builder, (string)value);
		}
		else if (value is bool)
		{
			builder.Append(Conversions.ToBoolean(value) ? "true" : "false");
		}
		else if (IsNumericType(valueType))
		{
			builder.Append(Convert.ToString(RuntimeHelpers.GetObjectValue(value), CultureInfo.InvariantCulture));
		}
		else if (value is DateTime)
		{
			WriteString(builder, ((DateTime)value).ToString("O", CultureInfo.InvariantCulture));
		}
		else if (value is IEnumerable && !(value is string))
		{
			WriteEnumerable(builder, (IEnumerable)value);
		}
		else
		{
			WriteObject(builder, RuntimeHelpers.GetObjectValue(value));
		}
	}

	private static void WriteObject(StringBuilder builder, object value)
	{
		builder.Append('{');
		bool first = true;
		PropertyInfo[] properties = value.GetType().GetProperties(BindingFlags.Instance | BindingFlags.Public);
		foreach (PropertyInfo property in properties)
		{
			if (property.CanRead && property.GetIndexParameters().Length == 0)
			{
				if (!first)
				{
					builder.Append(',');
				}
				first = false;
				WriteString(builder, property.Name);
				builder.Append(':');
				WriteValue(builder, RuntimeHelpers.GetObjectValue(property.GetValue(RuntimeHelpers.GetObjectValue(value), null)));
			}
		}
		builder.Append('}');
	}

	private static void WriteEnumerable(StringBuilder builder, IEnumerable values)
	{
		builder.Append('[');
		bool first = true;
		foreach (object value in values)
		{
			object item = RuntimeHelpers.GetObjectValue(value);
			if (!first)
			{
				builder.Append(',');
			}
			first = false;
			WriteValue(builder, RuntimeHelpers.GetObjectValue(item));
		}
		builder.Append(']');
	}

	private static void WriteString(StringBuilder builder, string value)
	{
		builder.Append('"');
		string text = value ?? string.Empty;
		foreach (char ch in text)
		{
			switch (ch)
			{
			case '"':
				builder.Append("\\\"");
				continue;
			case '\\':
				builder.Append("\\\\");
				continue;
			case '\b':
				builder.Append("\\b");
				continue;
			case '\f':
				builder.Append("\\f");
				continue;
			case '\n':
				builder.Append("\\n");
				continue;
			case '\r':
				builder.Append("\\r");
				continue;
			case '\t':
				builder.Append("\\t");
				continue;
			}
			if (ch < ' ')
			{
				builder.Append("\\u");
				int num = ch;
				builder.Append(num.ToString("x4", CultureInfo.InvariantCulture));
			}
			else
			{
				builder.Append(ch);
			}
		}
		builder.Append('"');
	}

	private static bool IsNumericType(Type valueType)
	{
		TypeCode typeCode = Type.GetTypeCode(valueType);
		if ((uint)(typeCode - 5) <= 10u)
		{
			return true;
		}
		return false;
	}
}
