using System;
using System.IO;
using System.Runtime.Serialization.Json;
using System.Text;
using Microsoft.VisualBasic.CompilerServices;

public sealed class DataContractJsonTextStore
{
	private DataContractJsonTextStore()
	{
	}

	public static T Load<T>(string json)
	{
		if (string.IsNullOrWhiteSpace(json))
		{
			throw new ArgumentException(FamilyBrowserLanguageService.Text("A JSON string is required.", "JSON 문자열이 필요합니다."), "json");
		}
		DataContractJsonSerializer serializer = new DataContractJsonSerializer(typeof(T));
		byte[] bytes = Encoding.UTF8.GetBytes(json);
		using MemoryStream stream = new MemoryStream(bytes);
		return Conversions.ToGenericParameter<T>(serializer.ReadObject(stream));
	}
}
