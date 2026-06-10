public sealed class SystemTypeIdentityService
{
	private SystemTypeIdentityService()
	{
	}

	public static string BuildKey(string typeClassName, string categoryName, string typeName)
	{
		return Normalize(typeClassName) + "|" + Normalize(categoryName) + "|" + Normalize(typeName);
	}

	public static string Normalize(string value)
	{
		if (value == null)
		{
			return string.Empty;
		}
		return value.Trim().ToLowerInvariant();
	}
}
