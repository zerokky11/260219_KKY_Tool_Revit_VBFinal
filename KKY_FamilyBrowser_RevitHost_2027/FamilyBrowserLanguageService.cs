using System;
using Microsoft.VisualBasic.CompilerServices;

public sealed class FamilyBrowserLanguageService
{
	private FamilyBrowserLanguageService()
	{
	}

	public static bool IsKorean()
	{
		try
		{
			return !string.Equals(FamilyBrowserUserSettingsStore.LoadLanguageCode(), "en", StringComparison.OrdinalIgnoreCase);
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ProjectData.ClearProjectError();
		}
		return true;
	}

	public static string Text(string englishText, string koreanText)
	{
		if (!IsKorean())
		{
			return englishText;
		}
		return koreanText;
	}
}
