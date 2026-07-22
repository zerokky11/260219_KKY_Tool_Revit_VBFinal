using System;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Net;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;
using System.Security.Cryptography;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

[DataContract]
internal sealed class FamilyBrowserProductUpdateManifest
{
	[DataMember(Name = "version")]
	public string Version { get; set; }

	[DataMember(Name = "url")]
	public string Url { get; set; }

	[DataMember(Name = "installerUrl")]
	public string InstallerUrl { get; set; }

	[DataMember(Name = "downloadUrl")]
	public string DownloadUrl { get; set; }

	[DataMember(Name = "homepageUrl")]
	public string HomepageUrl { get; set; }

	[DataMember(Name = "publishedAt")]
	public string PublishedAt { get; set; }

	[DataMember(Name = "notes")]
	public string Notes { get; set; }

	[DataMember(Name = "sha256")]
	public string Sha256 { get; set; }
}

internal sealed class FamilyBrowserProductUpdateResult
{
	public bool Success { get; set; }

	public bool HasUpdate { get; set; }

	public string CurrentVersion { get; set; }

	public string LatestVersion { get; set; }

	public string DownloadUrl { get; set; }

	public string HomepageUrl { get; set; }

	public string PublishedAt { get; set; }

	public string Notes { get; set; }

	public string InstallerSha256 { get; set; }

	public string ErrorMessage { get; set; }

	public FamilyBrowserProductUpdateResult()
	{
		CurrentVersion = FamilyBrowserProductUpdateService.CurrentProductVersion;
		LatestVersion = string.Empty;
		DownloadUrl = string.Empty;
		HomepageUrl = FamilyBrowserProductUpdateService.DefaultHomepageUrl;
		PublishedAt = string.Empty;
		Notes = string.Empty;
		InstallerSha256 = string.Empty;
		ErrorMessage = string.Empty;
	}
}

internal sealed class FamilyBrowserProductUpdateDownloadResult
{
	public bool Success { get; set; }

	public bool ReusedCachedFile { get; set; }

	public string InstallerPath { get; set; }

	public string Sha256 { get; set; }

	public long Bytes { get; set; }

	public string ErrorMessage { get; set; }

	public FamilyBrowserProductUpdateDownloadResult()
	{
		InstallerPath = string.Empty;
		Sha256 = string.Empty;
		ErrorMessage = string.Empty;
	}
}

internal static class FamilyBrowserProductUpdateService
{
	public const string CurrentProductVersion = "1.0.1";

	public const string DefaultFeedUrl = "https://update.zerokky.com/Release/family-browser/latest.json";

	public const string DefaultHomepageUrl = "https://update.zerokky.com/";

	public const string DefaultManualUrl = "https://update.zerokky.com/family-browser/index.html";

	private const int RequestTimeoutMilliseconds = 7000;

	private const int MaximumManifestCharacters = 262144;

	private const int InstallerDownloadTimeoutMilliseconds = 60000;

	private const long MinimumInstallerBytes = 131072L;

	private const long MaximumInstallerBytes = 268435456L;

	public static FamilyBrowserProductUpdateResult CheckForUpdates()
	{
		FamilyBrowserProductUpdateResult result = new FamilyBrowserProductUpdateResult();
		try
		{
			string json = DownloadManifestText(DefaultFeedUrl);
			FamilyBrowserProductUpdateManifest manifest = DeserializeManifest(json);
			if (manifest == null || string.IsNullOrWhiteSpace(manifest.Version))
			{
				throw new InvalidDataException("The Family Browser update feed does not contain a version value.");
			}

			result.LatestVersion = CleanVersionDisplay(manifest.Version);
			if (string.IsNullOrWhiteSpace(result.LatestVersion))
			{
				throw new InvalidDataException("The Family Browser update feed contains an invalid version value.");
			}

			result.HomepageUrl = ResolveHttpLocation(DefaultFeedUrl, manifest.HomepageUrl, DefaultHomepageUrl);
			string downloadLocation = FirstNonEmpty(manifest.DownloadUrl, manifest.InstallerUrl, manifest.Url);
			result.DownloadUrl = ResolveHttpLocation(DefaultFeedUrl, downloadLocation, string.Empty);
			result.PublishedAt = (manifest.PublishedAt ?? string.Empty).Trim();
			result.Notes = (manifest.Notes ?? string.Empty).Trim();
			result.InstallerSha256 = NormalizeSha256(manifest.Sha256);
			result.HasUpdate = CompareVersions(result.LatestVersion, result.CurrentVersion) > 0;
			if (result.HasUpdate)
			{
				Uri trustedInstallerUri;
				string trustError;
				if (!TryGetTrustedInstallerUri(result.DownloadUrl, out trustedInstallerUri, out trustError))
				{
					throw new InvalidDataException("The Family Browser update feed does not provide a trusted HTTPS installer URL. " + trustError);
				}
				if (!IsValidSha256(result.InstallerSha256))
				{
					throw new InvalidDataException("The Family Browser update feed does not provide a valid installer SHA-256 value.");
				}
				result.DownloadUrl = trustedInstallerUri.AbsoluteUri;
			}
			result.Success = true;
		}
		catch (Exception ex)
		{
			result.Success = false;
			result.ErrorMessage = ex.Message;
		}
		return result;
	}

	public static FamilyBrowserProductUpdateDownloadResult DownloadUpdateInstaller(FamilyBrowserProductUpdateResult update)
	{
		FamilyBrowserProductUpdateDownloadResult result = new FamilyBrowserProductUpdateDownloadResult();
		string temporaryPath = string.Empty;
		try
		{
			if (update == null || !update.HasUpdate)
			{
				throw new InvalidOperationException("No Family Browser update is available for download.");
			}

			Uri installerUri;
			string trustError;
			if (!TryGetTrustedInstallerUri(update.DownloadUrl, out installerUri, out trustError))
			{
				throw new InvalidDataException("The installer URL is not trusted. " + trustError);
			}

			string expectedSha256 = NormalizeSha256(update.InstallerSha256);
			if (!IsValidSha256(expectedSha256))
			{
				throw new InvalidDataException("The installer SHA-256 value is missing or invalid.");
			}

			string updateRoot = GetUpdateCacheRoot();
			string versionFolder = Path.Combine(updateRoot, SafePathSegment(update.LatestVersion, "update"));
			EnsurePathInsideRoot(updateRoot, versionFolder);
			Directory.CreateDirectory(versionFolder);

			string installerPath = BuildCachedInstallerPath(versionFolder, installerUri, expectedSha256);
			EnsurePathInsideRoot(updateRoot, installerPath);
			long cachedBytes;
			string cachedHash;
			string validationError;
			if (File.Exists(installerPath) && ValidateInstallerFile(installerPath, expectedSha256, out cachedBytes, out cachedHash, out validationError))
			{
				result.Success = true;
				result.ReusedCachedFile = true;
				result.InstallerPath = installerPath;
				result.Sha256 = cachedHash;
				result.Bytes = cachedBytes;
				return result;
			}
			if (File.Exists(installerPath))
			{
				File.Delete(installerPath);
			}

			temporaryPath = installerPath + "." + Guid.NewGuid().ToString("N") + ".download";
			EnsurePathInsideRoot(updateRoot, temporaryPath);
			HttpWebRequest request = CreateInstallerDownloadRequest(installerUri);
			using (HttpWebResponse response = (HttpWebResponse)request.GetResponse())
			{
				if ((int)response.StatusCode < 200 || (int)response.StatusCode >= 300)
				{
					throw new WebException("The Family Browser installer server returned HTTP " + ((int)response.StatusCode).ToString(CultureInfo.InvariantCulture) + ".");
				}
				Uri finalResponseUri;
				if (!TryGetTrustedInstallerUri(response.ResponseUri == null ? string.Empty : response.ResponseUri.AbsoluteUri, out finalResponseUri, out trustError))
				{
					throw new InvalidDataException("The installer download redirected to an untrusted location. " + trustError);
				}
				if (response.ContentLength > MaximumInstallerBytes)
				{
					throw new InvalidDataException("The Family Browser installer is larger than the allowed download limit.");
				}

				long totalBytes = 0L;
				byte[] buffer = new byte[65536];
				using (Stream source = response.GetResponseStream() ?? Stream.Null)
				using (FileStream destination = new FileStream(temporaryPath, FileMode.CreateNew, FileAccess.Write, FileShare.None))
				{
					int read;
					while ((read = source.Read(buffer, 0, buffer.Length)) > 0)
					{
						totalBytes += read;
						if (totalBytes > MaximumInstallerBytes)
						{
							throw new InvalidDataException("The Family Browser installer exceeded the allowed download limit.");
						}
						destination.Write(buffer, 0, read);
					}
					destination.Flush(true);
				}
			}

			long downloadedBytes;
			string downloadedHash;
			if (!ValidateInstallerFile(temporaryPath, expectedSha256, out downloadedBytes, out downloadedHash, out validationError))
			{
				throw new InvalidDataException("The downloaded Family Browser installer failed validation. " + validationError);
			}

			try
			{
				File.Move(temporaryPath, installerPath);
				temporaryPath = string.Empty;
			}
			catch (IOException)
			{
				long racedBytes;
				string racedHash;
				if (!File.Exists(installerPath) || !ValidateInstallerFile(installerPath, expectedSha256, out racedBytes, out racedHash, out validationError))
				{
					throw;
				}
				downloadedBytes = racedBytes;
				downloadedHash = racedHash;
			}

			try
			{
				File.WriteAllText(installerPath + ".sha256.txt", downloadedHash + "  " + Path.GetFileName(installerPath), Encoding.ASCII);
			}
			catch
			{
				// The sidecar is diagnostic only; the installer itself has already passed validation.
			}
			result.Success = true;
			result.InstallerPath = installerPath;
			result.Sha256 = downloadedHash;
			result.Bytes = downloadedBytes;
		}
		catch (Exception ex)
		{
			result.Success = false;
			result.ErrorMessage = ex.Message;
		}
		finally
		{
			if (!string.IsNullOrWhiteSpace(temporaryPath))
			{
				TryDeleteFile(temporaryPath);
			}
		}
		return result;
	}

	internal static bool IsTrustedInstallerUrlForAudit(string address)
	{
		Uri installerUri;
		string error;
		return TryGetTrustedInstallerUri(address, out installerUri, out error);
	}

	internal static bool ValidateInstallerFileForAudit(string path, string expectedSha256)
	{
		long bytes;
		string actualSha256;
		string error;
		return ValidateInstallerFile(path, expectedSha256, out bytes, out actualSha256, out error);
	}

	internal static bool ValidateDownloadedInstallerBeforeLaunch(string path, string expectedSha256)
	{
		long bytes;
		string actualSha256;
		string error;
		return ValidateInstallerFile(path, expectedSha256, out bytes, out actualSha256, out error);
	}

	private static string DownloadManifestText(string url)
	{
		HttpWebRequest request = WebRequest.CreateHttp(url);
		request.Method = "GET";
		request.Accept = "application/json,text/plain;q=0.9,*/*;q=0.5";
		request.UserAgent = "KKY-FamilyBrowser/" + CurrentProductVersion;
		request.AutomaticDecompression = DecompressionMethods.GZip | DecompressionMethods.Deflate;
		request.Timeout = RequestTimeoutMilliseconds;
		request.ReadWriteTimeout = RequestTimeoutMilliseconds;
		request.CachePolicy = new System.Net.Cache.RequestCachePolicy(System.Net.Cache.RequestCacheLevel.NoCacheNoStore);

		using (HttpWebResponse response = (HttpWebResponse)request.GetResponse())
		{
			if ((int)response.StatusCode < 200 || (int)response.StatusCode >= 300)
			{
				throw new WebException("The Family Browser update server returned HTTP " + ((int)response.StatusCode).ToString(CultureInfo.InvariantCulture) + ".");
			}
			using (Stream responseStream = response.GetResponseStream())
			using (StreamReader reader = new StreamReader(responseStream ?? Stream.Null, Encoding.UTF8, true))
			{
				string text = reader.ReadToEnd();
				if (text.Length > MaximumManifestCharacters)
				{
					throw new InvalidDataException("The Family Browser update feed is larger than expected.");
				}
				return text;
			}
		}
	}

	private static FamilyBrowserProductUpdateManifest DeserializeManifest(string json)
	{
		if (string.IsNullOrWhiteSpace(json))
		{
			throw new InvalidDataException("The Family Browser update feed is empty.");
		}
		DataContractJsonSerializer serializer = new DataContractJsonSerializer(typeof(FamilyBrowserProductUpdateManifest));
		using (MemoryStream stream = new MemoryStream(Encoding.UTF8.GetBytes(json)))
		{
			return serializer.ReadObject(stream) as FamilyBrowserProductUpdateManifest;
		}
	}

	private static string FirstNonEmpty(params string[] values)
	{
		foreach (string value in values ?? new string[0])
		{
			if (!string.IsNullOrWhiteSpace(value))
			{
				return value.Trim();
			}
		}
		return string.Empty;
	}

	private static string ResolveHttpLocation(string baseUrl, string candidate, string fallback)
	{
		string value = string.IsNullOrWhiteSpace(candidate) ? fallback : candidate.Trim();
		if (string.IsNullOrWhiteSpace(value))
		{
			return string.Empty;
		}

		Uri absolute;
		if (!Uri.TryCreate(value, UriKind.Absolute, out absolute))
		{
			Uri baseUri;
			if (!Uri.TryCreate(baseUrl, UriKind.Absolute, out baseUri) || !Uri.TryCreate(baseUri, value, out absolute))
			{
				return string.Empty;
			}
		}
		if (!string.Equals(absolute.Scheme, Uri.UriSchemeHttps, StringComparison.OrdinalIgnoreCase) &&
			!string.Equals(absolute.Scheme, Uri.UriSchemeHttp, StringComparison.OrdinalIgnoreCase))
		{
			return string.Empty;
		}
		return absolute.AbsoluteUri;
	}

	private static string CleanVersionDisplay(string value)
	{
		string clean = (value ?? string.Empty).Trim();
		if (clean.StartsWith("v", StringComparison.OrdinalIgnoreCase))
		{
			clean = clean.Substring(1).Trim();
		}
		int suffix = clean.IndexOfAny(new[] { '-', '+' });
		if (suffix >= 0)
		{
			clean = clean.Substring(0, suffix);
		}
		return clean.Trim();
	}

	private static int CompareVersions(string left, string right)
	{
		Version leftVersion = ParseComparableVersion(left);
		Version rightVersion = ParseComparableVersion(right);
		return leftVersion.CompareTo(rightVersion);
	}

	private static Version ParseComparableVersion(string value)
	{
		Version parsed;
		string clean = CleanVersionDisplay(value);
		if (!Version.TryParse(clean, out parsed))
		{
			throw new InvalidDataException("Invalid Family Browser version: " + clean);
		}
		return new Version(parsed.Major, parsed.Minor, Math.Max(parsed.Build, 0), Math.Max(parsed.Revision, 0));
	}

	private static HttpWebRequest CreateInstallerDownloadRequest(Uri installerUri)
	{
		HttpWebRequest request = WebRequest.CreateHttp(installerUri);
		request.Method = "GET";
		request.Accept = "application/octet-stream,application/x-msdownload;q=0.9,*/*;q=0.5";
		request.UserAgent = "KKY-FamilyBrowser-Updater/" + CurrentProductVersion;
		request.AutomaticDecompression = DecompressionMethods.GZip | DecompressionMethods.Deflate;
		request.AllowAutoRedirect = true;
		request.MaximumAutomaticRedirections = 3;
		request.Timeout = InstallerDownloadTimeoutMilliseconds;
		request.ReadWriteTimeout = InstallerDownloadTimeoutMilliseconds;
		request.CachePolicy = new System.Net.Cache.RequestCachePolicy(System.Net.Cache.RequestCacheLevel.NoCacheNoStore);
		return request;
	}

	private static bool TryGetTrustedInstallerUri(string address, out Uri installerUri, out string error)
	{
		installerUri = null;
		error = string.Empty;
		Uri candidate;
		if (!Uri.TryCreate(address, UriKind.Absolute, out candidate))
		{
			error = "The URL is missing or invalid.";
			return false;
		}
		if (!string.Equals(candidate.Scheme, Uri.UriSchemeHttps, StringComparison.OrdinalIgnoreCase))
		{
			error = "Only HTTPS downloads are allowed.";
			return false;
		}
		Uri feedUri = new Uri(DefaultFeedUrl);
		if (!string.Equals(candidate.DnsSafeHost, feedUri.DnsSafeHost, StringComparison.OrdinalIgnoreCase))
		{
			error = "The download host does not match the trusted update server.";
			return false;
		}
		if (!candidate.IsDefaultPort && candidate.Port != 443)
		{
			error = "The HTTPS download uses an unexpected port.";
			return false;
		}
		if (!string.IsNullOrWhiteSpace(candidate.UserInfo) || !string.IsNullOrWhiteSpace(candidate.Fragment))
		{
			error = "The installer URL contains unsupported authority or fragment data.";
			return false;
		}
		if (!candidate.AbsolutePath.EndsWith(".exe", StringComparison.OrdinalIgnoreCase))
		{
			error = "The trusted update artifact must be an EXE installer.";
			return false;
		}
		installerUri = candidate;
		return true;
	}

	private static string GetUpdateCacheRoot()
	{
		string localAppData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
		if (string.IsNullOrWhiteSpace(localAppData))
		{
			throw new InvalidOperationException("The local application data folder is unavailable.");
		}
		return Path.Combine(localAppData, "KKY", "FamilyBrowser", "Updates");
	}

	private static string BuildCachedInstallerPath(string versionFolder, Uri installerUri, string expectedSha256)
	{
		string rawName = Path.GetFileName(Uri.UnescapeDataString(installerUri.AbsolutePath));
		string baseName = Path.GetFileNameWithoutExtension(rawName);
		baseName = SafePathSegment(baseName, "KKY_FamilyBrowser_Update");
		return Path.Combine(versionFolder, baseName + "_" + expectedSha256.Substring(0, 12) + ".exe");
	}

	private static string SafePathSegment(string value, string fallback)
	{
		string clean = (value ?? string.Empty).Trim();
		foreach (char invalid in Path.GetInvalidFileNameChars())
		{
			clean = clean.Replace(invalid, '_');
		}
		clean = clean.Trim(' ', '.');
		return string.IsNullOrWhiteSpace(clean) ? fallback : clean;
	}

	private static void EnsurePathInsideRoot(string root, string path)
	{
		string rootFull = Path.GetFullPath(root).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar) + Path.DirectorySeparatorChar;
		string pathFull = Path.GetFullPath(path);
		if (!pathFull.StartsWith(rootFull, StringComparison.OrdinalIgnoreCase))
		{
			throw new InvalidOperationException("The update file path escaped the local Family Browser update cache.");
		}
	}

	private static bool ValidateInstallerFile(string path, string expectedSha256, out long bytes, out string actualSha256, out string error)
	{
		bytes = 0L;
		actualSha256 = string.Empty;
		error = string.Empty;
		try
		{
			if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
			{
				error = "The installer file does not exist.";
				return false;
			}
			FileInfo info = new FileInfo(path);
			bytes = info.Length;
			if (bytes < MinimumInstallerBytes || bytes > MaximumInstallerBytes)
			{
				error = "The installer size is outside the allowed range.";
				return false;
			}
			using (FileStream stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read))
			{
				if (stream.ReadByte() != 0x4D || stream.ReadByte() != 0x5A)
				{
					error = "The installer does not have a valid PE header.";
					return false;
				}
			}
			actualSha256 = ComputeFileSha256(path);
			string expected = NormalizeSha256(expectedSha256);
			if (!IsValidSha256(expected) || !FixedTimeEquals(actualSha256, expected))
			{
				error = "The installer SHA-256 does not match the release feed value.";
				return false;
			}
			return true;
		}
		catch (Exception ex)
		{
			error = ex.Message;
			return false;
		}
	}

	private static string ComputeFileSha256(string path)
	{
		using (FileStream stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read))
		using (SHA256 sha = SHA256.Create())
		{
			return BitConverter.ToString(sha.ComputeHash(stream)).Replace("-", string.Empty);
		}
	}

	private static string NormalizeSha256(string value)
	{
		string clean = (value ?? string.Empty).Trim();
		if (clean.StartsWith("sha256:", StringComparison.OrdinalIgnoreCase))
		{
			clean = clean.Substring(7).Trim();
		}
		return clean.Replace(" ", string.Empty).ToUpperInvariant();
	}

	private static bool IsValidSha256(string value)
	{
		string clean = NormalizeSha256(value);
		if (clean.Length != 64)
		{
			return false;
		}
		for (int index = 0; index < clean.Length; index++)
		{
			char character = clean[index];
			if (!((character >= '0' && character <= '9') || (character >= 'A' && character <= 'F')))
			{
				return false;
			}
		}
		return true;
	}

	private static bool FixedTimeEquals(string left, string right)
	{
		string a = NormalizeSha256(left);
		string b = NormalizeSha256(right);
		int difference = a.Length ^ b.Length;
		int length = Math.Max(a.Length, b.Length);
		for (int index = 0; index < length; index++)
		{
			char leftCharacter = index < a.Length ? a[index] : (char)0;
			char rightCharacter = index < b.Length ? b[index] : (char)0;
			difference |= leftCharacter ^ rightCharacter;
		}
		return difference == 0;
	}

	private static void TryDeleteFile(string path)
	{
		try
		{
			if (File.Exists(path))
			{
				File.Delete(path);
			}
		}
		catch
		{
		}
	}
}

public partial class FamilyBrowserDashboardHtmlForm
{
	private bool _productUpdateCheckPending;

	private bool _productUpdateDownloadPending;

	private void CheckFamilyBrowserProductUpdate()
	{
		if (_productUpdateCheckPending || _productUpdateDownloadPending)
		{
			_statusMessage = T("Update check is already running.", "업데이트 확인이 이미 진행 중입니다.");
			SetDashboardElementText("dashboardStatusText", _statusMessage);
			return;
		}

		_productUpdateCheckPending = true;
		_statusMessage = T("Checking the Family Browser version...", "Family Browser 버전을 확인하는 중입니다...");
		SetDashboardElementText("dashboardStatusText", _statusMessage);
		WriteDashboardRuntimeDiagnostic("product-update-check-start", -1, -1L);

		Task.Factory.StartNew(
			() => FamilyBrowserProductUpdateService.CheckForUpdates(),
			CancellationToken.None,
			TaskCreationOptions.DenyChildAttach,
			TaskScheduler.Default).ContinueWith(completed =>
			{
				if (IsDisposed)
				{
					return;
				}
				try
				{
					BeginInvoke((Action)delegate
					{
						FamilyBrowserProductUpdateResult result = completed.Status == TaskStatus.RanToCompletion
							? completed.Result
							: new FamilyBrowserProductUpdateResult
							{
								Success = false,
								ErrorMessage = completed.Exception == null ? "Unknown update-check failure." : completed.Exception.GetBaseException().Message
							};
						CompleteFamilyBrowserProductUpdateCheck(result);
					});
				}
				catch (Exception ex)
				{
					WriteDashboardRuntimeDiagnostic("product-update-check-ui-failed:" + ex.Message, -1, -1L);
				}
			}, TaskScheduler.Default);
	}

	private void CompleteFamilyBrowserProductUpdateCheck(FamilyBrowserProductUpdateResult result)
	{
		_productUpdateCheckPending = false;
		string caption = T("Family Browser Update", "Family Browser 업데이트 확인");
		if (result == null || !result.Success)
		{
			string reason = result == null ? T("No result was returned.", "확인 결과를 받지 못했습니다.") : (result.ErrorMessage ?? string.Empty);
			_statusMessage = T("Update check failed.", "업데이트 확인에 실패했습니다.");
			SetDashboardElementText("dashboardStatusText", _statusMessage);
			WriteDashboardRuntimeDiagnostic("product-update-check-failed:" + reason, -1, -1L);
			FamilyBrowserResultDialog.Show(
				this,
				caption,
				T("The update server could not be checked.", "업데이트 서버를 확인하지 못했습니다.") + "\r\n\r\n" +
				T("Current version", "현재 버전") + ": " + FamilyBrowserProductUpdateService.CurrentProductVersion + "\r\n" +
				T("Details", "상세") + ": " + reason + "\r\n\r\n" +
				T("Check the network connection or open Homepage from the Support group.", "네트워크 연결을 확인하거나 지원 그룹의 홈페이지를 열어 확인하세요."),
				MessageBoxIcon.Exclamation);
			return;
		}

		string details = BuildFamilyBrowserProductUpdateDetails(result);
		if (result.HasUpdate)
		{
			_statusMessage = T("A Family Browser update is available: ", "Family Browser 새 버전 사용 가능: ") + result.LatestVersion;
			SetDashboardElementText("dashboardStatusText", _statusMessage);
			WriteDashboardRuntimeDiagnostic("product-update-check-update-available:" + result.LatestVersion, -1, -1L);
			bool downloadUpdate = FamilyBrowserResultDialog.Confirm(
				this,
				caption,
				T("A newer Family Browser version is available.", "새로운 Family Browser 버전이 있습니다."),
				details + "\r\n\r\n" + T("The installer will be downloaded from the trusted update server and verified before it can run.", "설치파일은 신뢰된 업데이트 서버에서 다운로드한 뒤 검증을 통과해야 실행할 수 있습니다."),
				T("Download Update", "업데이트 다운로드"),
				T("Close", "닫기"),
				false);
			if (downloadUpdate)
			{
				BeginFamilyBrowserProductUpdateDownload(result);
			}
			return;
		}

		_statusMessage = T("Family Browser is up to date: ", "Family Browser가 최신 버전입니다: ") + result.CurrentVersion;
		SetDashboardElementText("dashboardStatusText", _statusMessage);
		WriteDashboardRuntimeDiagnostic("product-update-check-current:" + result.CurrentVersion, -1, -1L);
		FamilyBrowserResultDialog.Show(
			this,
			caption,
			T("You are using the latest Family Browser version.", "현재 최신 Family Browser 버전을 사용 중입니다.") + "\r\n\r\n" + details,
			MessageBoxIcon.Information);
	}

	private void BeginFamilyBrowserProductUpdateDownload(FamilyBrowserProductUpdateResult update)
	{
		if (_productUpdateDownloadPending)
		{
			return;
		}
		_productUpdateDownloadPending = true;
		_statusMessage = T("Downloading and verifying the Family Browser update...", "Family Browser 업데이트를 다운로드하고 검증하는 중입니다...");
		SetDashboardElementText("dashboardStatusText", _statusMessage);
		WriteDashboardRuntimeDiagnostic("product-update-download-start:" + (update == null ? string.Empty : update.LatestVersion), -1, -1L);

		Task.Factory.StartNew(
			() => FamilyBrowserProductUpdateService.DownloadUpdateInstaller(update),
			CancellationToken.None,
			TaskCreationOptions.DenyChildAttach,
			TaskScheduler.Default).ContinueWith(completed =>
			{
				if (IsDisposed)
				{
					return;
				}
				try
				{
					BeginInvoke((Action)delegate
					{
						FamilyBrowserProductUpdateDownloadResult download = completed.Status == TaskStatus.RanToCompletion
							? completed.Result
							: new FamilyBrowserProductUpdateDownloadResult
							{
								Success = false,
								ErrorMessage = completed.Exception == null ? "Unknown update-download failure." : completed.Exception.GetBaseException().Message
							};
						CompleteFamilyBrowserProductUpdateDownload(update, download);
					});
				}
				catch (Exception ex)
				{
					WriteDashboardRuntimeDiagnostic("product-update-download-ui-failed:" + ex.Message, -1, -1L);
				}
			}, TaskScheduler.Default);
	}

	private void CompleteFamilyBrowserProductUpdateDownload(FamilyBrowserProductUpdateResult update, FamilyBrowserProductUpdateDownloadResult download)
	{
		_productUpdateDownloadPending = false;
		string caption = T("Family Browser Update", "Family Browser 업데이트");
		if (download == null || !download.Success)
		{
			string reason = download == null ? T("No download result was returned.", "다운로드 결과를 받지 못했습니다.") : (download.ErrorMessage ?? string.Empty);
			_statusMessage = T("Update download or verification failed.", "업데이트 다운로드 또는 검증에 실패했습니다.");
			SetDashboardElementText("dashboardStatusText", _statusMessage);
			WriteDashboardRuntimeDiagnostic("product-update-download-failed:" + reason, -1, -1L);
			bool openHomepage = FamilyBrowserResultDialog.Confirm(
				this,
				caption,
				_statusMessage,
				reason,
				T("Open Homepage", "홈페이지 열기"),
				T("Close", "닫기"),
				false);
			if (openHomepage)
			{
				OpenFamilyBrowserSupportUri(update == null ? FamilyBrowserProductUpdateService.DefaultHomepageUrl : update.HomepageUrl, T("Homepage", "홈페이지"));
			}
			return;
		}

		_statusMessage = T("The Family Browser update installer is ready.", "Family Browser 업데이트 설치파일이 준비되었습니다.");
		SetDashboardElementText("dashboardStatusText", _statusMessage);
		WriteDashboardRuntimeDiagnostic("product-update-download-verified:" + (update == null ? string.Empty : update.LatestVersion) + ":" + download.Sha256, -1, download.Bytes);

		StringBuilder details = new StringBuilder();
		details.AppendLine(T("Version", "버전") + ": " + (update == null ? "-" : update.LatestVersion));
		details.AppendLine(T("Size", "크기") + ": " + Math.Round(download.Bytes / 1048576.0, 2).ToString(CultureInfo.InvariantCulture) + " MB");
		details.AppendLine("SHA-256: " + download.Sha256);
		details.AppendLine();
		details.AppendLine(T("Save or synchronize every open project before starting the installer.", "설치 시작 전에 열려 있는 모든 프로젝트를 저장하거나 동기화하세요."));
		details.AppendLine(T("The installer requires administrator approval and Revit must close before files can be replaced.", "설치에는 관리자 승인이 필요하며 파일을 교체하려면 Revit을 종료해야 합니다."));

		bool startInstaller = FamilyBrowserResultDialog.Confirm(
			this,
			caption,
			T("The downloaded installer passed integrity verification.", "다운로드한 설치파일이 무결성 검증을 통과했습니다."),
			details.ToString(),
			T("Start Installer", "설치 시작"),
			T("Later", "나중에"),
			false);
		if (!startInstaller)
		{
			return;
		}

		StartFamilyBrowserUpdateInstaller(download);
	}

	private void StartFamilyBrowserUpdateInstaller(FamilyBrowserProductUpdateDownloadResult download)
	{
		string caption = T("Family Browser Update", "Family Browser 업데이트");
		try
		{
			if (download == null || !download.Success || !FamilyBrowserProductUpdateService.ValidateDownloadedInstallerBeforeLaunch(download.InstallerPath, download.Sha256))
			{
				throw new InvalidDataException(T("The installer changed after download verification, so it was not executed.", "다운로드 검증 후 설치파일이 변경되어 실행하지 않았습니다."));
			}
			ProcessStartInfo startInfo = new ProcessStartInfo
			{
				FileName = download.InstallerPath,
				WorkingDirectory = Path.GetDirectoryName(download.InstallerPath),
				UseShellExecute = true,
				Verb = "runas"
			};
			Process.Start(startInfo);
			_statusMessage = T("The verified update installer was started.", "검증된 업데이트 설치파일을 실행했습니다.");
			SetDashboardElementText("dashboardStatusText", _statusMessage);
			WriteDashboardRuntimeDiagnostic("product-update-installer-started:" + download.Sha256, -1, download.Bytes);
			FamilyBrowserResultDialog.Show(
				this,
				caption,
				T("Complete the installer after saving and closing every Revit window. Start Revit again after installation finishes.", "열려 있는 Revit 작업을 저장하고 모든 Revit 창을 닫은 뒤 설치를 완료하세요. 설치가 끝나면 Revit을 다시 실행하세요."),
				MessageBoxIcon.Information);
		}
		catch (Exception ex)
		{
			_statusMessage = T("The update installer could not be started.", "업데이트 설치파일을 실행하지 못했습니다.");
			SetDashboardElementText("dashboardStatusText", _statusMessage);
			WriteDashboardRuntimeDiagnostic("product-update-installer-start-failed:" + ex.Message, -1, -1L);
			FamilyBrowserResultDialog.Show(this, caption, _statusMessage + "\r\n\r\n" + ex.Message, MessageBoxIcon.Exclamation);
		}
	}

	private string BuildFamilyBrowserProductUpdateDetails(FamilyBrowserProductUpdateResult result)
	{
		StringBuilder details = new StringBuilder();
		details.AppendLine(T("Current version", "현재 버전") + ": " + (result.CurrentVersion ?? "-"));
		details.AppendLine(T("Latest version", "최신 버전") + ": " + (result.LatestVersion ?? "-"));
		if (!string.IsNullOrWhiteSpace(result.PublishedAt))
		{
			details.AppendLine(T("Published", "게시일") + ": " + result.PublishedAt);
		}
		if (!string.IsNullOrWhiteSpace(result.Notes))
		{
			details.AppendLine();
			details.AppendLine(T("Release notes", "업데이트 내용"));
			details.AppendLine(result.Notes);
		}
		return details.ToString().Trim();
	}

	private void OpenFamilyBrowserHomepage()
	{
		OpenFamilyBrowserSupportUri(FamilyBrowserProductUpdateService.DefaultHomepageUrl, T("Homepage", "홈페이지"));
	}

	private void OpenFamilyBrowserManual()
	{
		OpenFamilyBrowserSupportUri(FamilyBrowserProductUpdateService.DefaultManualUrl, T("Manual", "매뉴얼"));
	}

	private void OpenFamilyBrowserSupportUri(string address, string label)
	{
		try
		{
			Uri uri;
			if (!Uri.TryCreate(address, UriKind.Absolute, out uri) ||
				(!string.Equals(uri.Scheme, Uri.UriSchemeHttps, StringComparison.OrdinalIgnoreCase) &&
				 !string.Equals(uri.Scheme, Uri.UriSchemeHttp, StringComparison.OrdinalIgnoreCase)))
			{
				throw new InvalidOperationException(T("The support link is invalid.", "지원 링크가 올바르지 않습니다."));
			}
			ProcessStartInfo startInfo = new ProcessStartInfo
			{
				FileName = uri.AbsoluteUri,
				UseShellExecute = true
			};
			Process.Start(startInfo);
			_statusMessage = label + T(" opened in the default browser.", "를 기본 브라우저에서 열었습니다.");
			SetDashboardElementText("dashboardStatusText", _statusMessage);
			WriteDashboardRuntimeDiagnostic("support-link-open:" + uri.AbsoluteUri, -1, -1L);
		}
		catch (Exception ex)
		{
			_statusMessage = label + T(" could not be opened.", "를 열지 못했습니다.");
			SetDashboardElementText("dashboardStatusText", _statusMessage);
			WriteDashboardRuntimeDiagnostic("support-link-open-failed:" + ex.Message, -1, -1L);
			FamilyBrowserResultDialog.Show(this, label, _statusMessage + "\r\n\r\n" + ex.Message, MessageBoxIcon.Exclamation);
		}
	}
}
