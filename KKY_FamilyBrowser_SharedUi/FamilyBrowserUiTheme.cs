using System;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;

public enum FamilyBrowserUiTheme
{
	Light,
	Dark
}

public interface IFamilyBrowserThemeAware
{
	void ApplyTheme(FamilyBrowserUiTheme theme);
}

public static class FamilyBrowserUiThemeService
{
	private static readonly object SyncRoot = new object();

	private static string _themeCss = string.Empty;

	private static string _brandMarkDataUri = string.Empty;

	public static FamilyBrowserUiTheme Load()
	{
		return FamilyBrowserUiTheme.Light;
	}

	public static bool Save(FamilyBrowserUiTheme theme, out string error)
	{
		error = string.Empty;
		return true;
	}

	public static FamilyBrowserUiTheme Parse(string value)
	{
		return FamilyBrowserUiTheme.Light;
	}

	public static string Code(FamilyBrowserUiTheme theme)
	{
		return "light";
	}

	public static string BodyClass(FamilyBrowserUiTheme theme)
	{
		return "theme-" + Code(theme);
	}

	public static string ThemeCss()
	{
		lock (SyncRoot)
		{
			if (!string.IsNullOrEmpty(_themeCss))
			{
				return _themeCss;
			}
			_themeCss = ReadEmbeddedText("family-browser-theme.css");
			return _themeCss;
		}
	}

	public static string BrandMarkDataUri()
	{
		lock (SyncRoot)
		{
			if (!string.IsNullOrEmpty(_brandMarkDataUri))
			{
				return _brandMarkDataUri;
			}
			byte[] bytes = ReadEmbeddedBytes("kky-tool-mark-24.png");
			_brandMarkDataUri = bytes.Length == 0 ? string.Empty : "data:image/png;base64," + Convert.ToBase64String(bytes);
			return _brandMarkDataUri;
		}
	}

	private static string ReadEmbeddedText(string fileName)
	{
		byte[] bytes = ReadEmbeddedBytes(fileName);
		return bytes.Length == 0 ? string.Empty : Encoding.UTF8.GetString(bytes);
	}

	private static byte[] ReadEmbeddedBytes(string fileName)
	{
		try
		{
			Assembly assembly = typeof(FamilyBrowserUiThemeService).Assembly;
			string resourceName = assembly.GetManifestResourceNames().FirstOrDefault(name => name.EndsWith(fileName, StringComparison.OrdinalIgnoreCase));
			if (string.IsNullOrWhiteSpace(resourceName))
			{
				return new byte[0];
			}
			using (Stream stream = assembly.GetManifestResourceStream(resourceName))
			{
				if (stream == null)
				{
					return new byte[0];
				}
				using (MemoryStream buffer = new MemoryStream())
				{
					stream.CopyTo(buffer);
					return buffer.ToArray();
				}
			}
		}
		catch (Exception ex)
		{
			WriteDiagnostic("resource/" + fileName, ex.Message);
			return new byte[0];
		}
	}

	private static void WriteDiagnostic(string stage, string detail)
	{
		try
		{
			string folder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "KKY", "FamilyBrowser", "Diagnostics");
			Directory.CreateDirectory(folder);
			string line = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff", CultureInfo.InvariantCulture) + " theme " + (stage ?? string.Empty) + " " + (detail ?? string.Empty) + Environment.NewLine;
			File.AppendAllText(Path.Combine(folder, "ui-theme.log"), line, Encoding.UTF8);
		}
		catch
		{
		}
	}
}
