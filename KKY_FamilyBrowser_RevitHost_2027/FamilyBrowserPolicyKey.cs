public sealed class FamilyBrowserPolicyKey
{
	private FamilyBrowserPolicyKey()
	{
	}

	public static string Normalize(string value)
	{
		if (value == null)
		{
			return string.Empty;
		}
		return value.Trim().ToLowerInvariant().Replace(' ', '-');
	}
}
