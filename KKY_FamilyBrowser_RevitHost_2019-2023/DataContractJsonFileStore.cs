using System;
using System.IO;
using System.Runtime.Serialization.Json;
using System.Text;
using Microsoft.VisualBasic.CompilerServices;

public sealed class DataContractJsonFileStore
{
	private DataContractJsonFileStore()
	{
	}

	public static T Load<T>(string path)
	{
		if (string.IsNullOrWhiteSpace(path))
		{
			throw new ArgumentException(FamilyBrowserLanguageService.Text("A JSON path is required.", "JSON 경로가 필요합니다."), "path");
		}
		if (!File.Exists(path))
		{
			throw new FileNotFoundException(FamilyBrowserLanguageService.Text("JSON file was not found.", "JSON 파일을 찾지 못했습니다."), path);
		}
		string json;
		using (StreamReader reader = new StreamReader(path, Encoding.UTF8, detectEncodingFromByteOrderMarks: true))
		{
			json = reader.ReadToEnd();
		}
		DataContractJsonSerializer serializer = new DataContractJsonSerializer(typeof(T));
		byte[] bytes = Encoding.UTF8.GetBytes(json ?? string.Empty);
		using MemoryStream stream = new MemoryStream(bytes);
		return Conversions.ToGenericParameter<T>(serializer.ReadObject(stream));
	}
}
