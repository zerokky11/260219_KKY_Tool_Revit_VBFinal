using System;
using System.IO;
using System.Reflection;
using System.Windows.Media;
using System.Windows.Media.Imaging;

public static class FamilyBrowserRibbonIcon
{
	private const string SmallResourceName = "KKY.FamilyBrowser.RibbonAssets.family-browser-ribbon-16.png";

	private const string LargeResourceName = "KKY.FamilyBrowser.RibbonAssets.family-browser-ribbon-32.png";

	public static ImageSource LoadSmall()
	{
		return Load(SmallResourceName, 16);
	}

	public static ImageSource LoadLarge()
	{
		return Load(LargeResourceName, 32);
	}

	private static ImageSource Load(string resourceName, int decodePixelWidth)
	{
		try
		{
			Assembly assembly = typeof(FamilyBrowserRibbonIcon).Assembly;
			using (Stream stream = assembly.GetManifestResourceStream(resourceName))
			{
				if (stream == null)
				{
					return null;
				}
				BitmapImage image = new BitmapImage();
				image.BeginInit();
				image.CacheOption = BitmapCacheOption.OnLoad;
				image.DecodePixelWidth = decodePixelWidth;
				image.StreamSource = stream;
				image.EndInit();
				image.Freeze();
				return image;
			}
		}
		catch (Exception)
		{
			return null;
		}
	}
}
