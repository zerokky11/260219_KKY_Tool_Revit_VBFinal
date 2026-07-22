using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.IO.Compression;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows.Forms;

internal static class Program
{
	private const uint SemFailCriticalErrors = 0x0001;

	private const uint SemNoGpFaultErrorBox = 0x0002;

	private const uint SemNoOpenFileErrorBox = 0x8000;

	private const uint LoadLibrarySearchDefaultDirs = 0x00001000;

	private const uint LoadLibrarySearchUserDirs = 0x00000400;

	private static readonly Regex HangulRegex = new Regex("[\u3131-\u318E\uAC00-\uD7A3]", RegexOptions.Compiled);

	private static readonly Regex LatinRegex = new Regex("[A-Za-z]", RegexOptions.Compiled);

	private static readonly string[] KoreanModeDisallowedEnglishPhrases = new[]
	{
		"Home",
		"Library",
		"Families",
		"System Types",
		"Workflow",
		"Requests",
		"Operations",
		"Model Check",
		"Unregistered",
		"Permissions",
		"Standard Management",
		"Debug Log",
		"Refresh",
		"Load Available",
		"Update Available",
		"Permission",
		"Tracking",
		"Readiness",
		"Detailed Filter",
		"visible",
		"Select",
		"Clear",
		"Detail",
		"Status",
		"Category",
		"Current screen",
		"not registered",
		"No active",
		"Open Folder",
		"Save",
		"Export",
		"Request created"
	};

	[DllImport("kernel32.dll")]
	private static extern uint SetErrorMode(uint uMode);

	[DllImport("kernel32.dll")]
	private static extern void ExitProcess(uint uExitCode);

	[DllImport("kernel32.dll", SetLastError = true)]
	private static extern bool SetDefaultDllDirectories(uint directoryFlags);

	[DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
	private static extern IntPtr AddDllDirectory(string newDirectory);

	[DllImport("user32.dll", SetLastError = true)]
	private static extern bool PrintWindow(IntPtr windowHandle, IntPtr deviceContext, uint flags);

	[DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
	private static extern bool CreateHardLink(string fileName, string existingFileName, IntPtr securityAttributes);

	private sealed class Clickable
	{
		public string Tag { get; set; }

		public string Text { get; set; }

		public string Href { get; set; }

		public string OnClick { get; set; }

		public string ClassName { get; set; }

		public int Index { get; set; }

		public HtmlElement Element { get; set; }
	}

	private sealed class AuditResult
	{
		public string Scenario { get; set; }

		public string HostAssembly { get; set; }

		public int ClickableCount { get; set; }

		public int BrowserClickCount { get; set; }

		public int HostActionCandidateCount { get; set; }

		public long HtmlRenderMilliseconds { get; set; }

		public long ColdHtmlRenderMilliseconds { get; set; }

		public long StartupShellRenderMilliseconds { get; set; }

		public long StartupShellLength { get; set; }

		public long HtmlLength { get; set; }

		public long DocumentLoadMilliseconds { get; set; }

		public long DashboardReadyMilliseconds { get; set; }

		public long FilterMilliseconds { get; set; }

		public long ThemeToggleMilliseconds { get; set; }

		public int DataRowCount { get; set; }

		public int DomRowCount { get; set; }

		public int VisibleRowCount { get; set; }

		public long CacheBytes { get; set; }

		public long CacheSaveMilliseconds { get; set; }

		public long CacheColdLoadMilliseconds { get; set; }

		public long CacheWarmLoadMilliseconds { get; set; }

		public long CacheOfflineLoadMilliseconds { get; set; }

		public readonly List<string> HostActions = new List<string>();

		public readonly List<string> Warnings = new List<string>();

		public readonly List<string> Failures = new List<string>();
	}

	[STAThread]
	private static int Main(string[] args)
	{
		SetErrorMode(SemFailCriticalErrors | SemNoGpFaultErrorBox | SemNoOpenFileErrorBox);
		Dictionary<string, string> options = ParseArgs(args);
		AuditResult result = new AuditResult
		{
			Scenario = GetOption(options, "scenario", "audit-default"),
			HostAssembly = GetOption(options, "assembly", string.Empty)
		};

		try
		{
			if (string.IsNullOrWhiteSpace(result.HostAssembly) || !File.Exists(result.HostAssembly))
			{
				throw new FileNotFoundException("Host assembly is missing.", result.HostAssembly);
			}

			List<string> dependencyDirs = BuildDependencyDirs(options, result.HostAssembly);
			ExtendProcessPath(dependencyDirs);
			RegisterAssemblyResolver(dependencyDirs);
			PreloadKnownDependencies(dependencyDirs, result);
			// Revit loads the add-in assembly before the user opens the browser command.
			// Keep host DLL load/JIT discovery out of the form-shell responsiveness metric.
			Assembly.LoadFrom(result.HostAssembly);
			if (GetBool(options, "revisionAudit", false))
			{
				RunStandardRevisionPrimitiveAudit(result.HostAssembly, result);
			}
			RunFileGuardPolicyAudit(result.HostAssembly, result);
			RunManagedFolderSetupAudit(result.HostAssembly, result);
			RunFamilyEditDialogGuardAudit(result.HostAssembly, result);
			RunMeasurementUnitPreferenceAudit(result.HostAssembly, result);
			RunProductUpdatePrimitiveAudit(result.HostAssembly, result);
			bool performanceMode = GetBool(options, "performanceMode", false);
			if (performanceMode)
			{
				Stopwatch startupShellSw = Stopwatch.StartNew();
				string startupShellHtml = RenderStartupShellHtml(options, result.HostAssembly);
				startupShellSw.Stop();
				result.StartupShellRenderMilliseconds = startupShellSw.ElapsedMilliseconds;
				result.StartupShellLength = startupShellHtml == null ? 0L : startupShellHtml.Length;
				if (string.IsNullOrWhiteSpace(startupShellHtml) || startupShellHtml.IndexOf("<html", StringComparison.OrdinalIgnoreCase) < 0)
				{
					result.Failures.Add("Rendered startup shell HTML is empty or invalid.");
				}
			}
			if (GetBool(options, "cacheAudit", false))
			{
				RunSyntheticCacheAudit(options, result);
			}
			string renderMode = GetOption(options, "renderMode", "dashboard");
			bool messageAudit = string.Equals(renderMode, "message", StringComparison.OrdinalIgnoreCase);
			Stopwatch renderSw = Stopwatch.StartNew();
			string html = messageAudit ? RenderMessageDialogHtml(options, result.HostAssembly) : RenderScenarioHtml(options, result.HostAssembly);
			renderSw.Stop();
			if (performanceMode)
			{
				result.ColdHtmlRenderMilliseconds = renderSw.ElapsedMilliseconds;
				renderSw.Restart();
				html = messageAudit ? RenderMessageDialogHtml(options, result.HostAssembly) : RenderScenarioHtml(options, result.HostAssembly);
				renderSw.Stop();
			}
			result.HtmlRenderMilliseconds = renderSw.ElapsedMilliseconds;
			result.HtmlLength = html == null ? 0L : html.Length;
			if (string.IsNullOrWhiteSpace(html) || html.IndexOf("<html", StringComparison.OrdinalIgnoreCase) < 0)
			{
				result.Failures.Add("Rendered HTML is empty or invalid.");
			}

			string htmlPath = GetOption(options, "htmlOut", string.Empty);
			if (!string.IsNullOrWhiteSpace(htmlPath))
			{
				Directory.CreateDirectory(Path.GetDirectoryName(Path.GetFullPath(htmlPath)));
				File.WriteAllText(htmlPath, html, new UTF8Encoding(false));
			}

			if (messageAudit)
			{
				RunMessageBodyAudit(html, options, result);
			}
			else
			{
				RunWebBrowserAudit(html, options, result);
			}
		}
		catch (Exception ex)
		{
			result.Failures.Add(DescribeException(ex));
		}

		string jsonPath = GetOption(options, "jsonOut", string.Empty);
		if (!string.IsNullOrWhiteSpace(jsonPath))
		{
			Directory.CreateDirectory(Path.GetDirectoryName(Path.GetFullPath(jsonPath)));
			File.WriteAllText(jsonPath, BuildJson(result), new UTF8Encoding(false));
		}

		Console.WriteLine(BuildConsoleSummary(result));
		int exitCode = result.Failures.Count == 0 ? 0 : 1;
		ExitProcess((uint)exitCode);
		return exitCode;
	}

	private static void RunStandardRevisionPrimitiveAudit(string hostAssembly, AuditResult result)
	{
		string fixtureRoot = Path.Combine(Path.GetTempPath(), "kkyfb-standard-revision-" + Guid.NewGuid().ToString("N"));
		try
		{
			Assembly assembly = Assembly.LoadFrom(hostAssembly);
			Type revisionType = assembly.GetType("FamilyBrowserStandardRevisionService", true);
			Type identityType = assembly.GetType("FamilyBrowserPathIdentityService", true);
			MethodInfo computeHash = revisionType.GetMethod("ComputeRevisionHash", BindingFlags.Public | BindingFlags.Static);
			MethodInfo comparableIdentity = identityType.GetMethod("GetComparableIdentity", BindingFlags.Public | BindingFlags.Static);
			if (computeHash == null || comparableIdentity == null)
			{
				result.Failures.Add("Standard revision primitive audit: required public methods are missing.");
				return;
			}

			Directory.CreateDirectory(fixtureRoot);
			string sourcePath = Path.Combine(fixtureRoot, "standard-source.rvt");
			byte[] content = new byte[4 * 1024 * 1024 + 17];
			for (int i = 0; i < content.Length; i++)
			{
				content[i] = (byte)((i * 31 + 17) % 251);
			}
			File.WriteAllBytes(sourcePath, content);
			DateTime fixedStamp = DateTime.UtcNow.AddMinutes(-10.0);
			File.SetLastWriteTimeUtc(sourcePath, fixedStamp);
			long fixedLength = new FileInfo(sourcePath).Length;
			string beforeHash = Convert.ToString(computeHash.Invoke(null, new object[] { sourcePath }));

			using (FileStream stream = new FileStream(sourcePath, FileMode.Open, FileAccess.ReadWrite, FileShare.ReadWrite | FileShare.Delete))
			{
				stream.Position = stream.Length / 2L;
				int current = stream.ReadByte();
				stream.Position--;
				stream.WriteByte((byte)((current + 1) % 256));
				stream.Flush(true);
			}
			File.SetLastWriteTimeUtc(sourcePath, fixedStamp);
			FileInfo modified = new FileInfo(sourcePath);
			string afterHash = Convert.ToString(computeHash.Invoke(null, new object[] { sourcePath }));
			if (fixedLength != modified.Length || Math.Abs((modified.LastWriteTimeUtc - fixedStamp).TotalSeconds) > 1.0 || string.Equals(beforeHash, afterHash, StringComparison.OrdinalIgnoreCase))
			{
				result.Failures.Add("Standard revision primitive audit: same-stamp/same-size content mutation was not detected by the sampled revision hash.");
			}

			string aliasPath = Path.Combine(fixtureRoot, "standard-source-alias.rvt");
			if (!CreateHardLink(aliasPath, sourcePath, IntPtr.Zero))
			{
				result.Failures.Add("Standard revision primitive audit: hard-link alias fixture could not be created. Win32=" + Marshal.GetLastWin32Error().ToString());
			}
			else
			{
				string sourceIdentity = Convert.ToString(comparableIdentity.Invoke(null, new object[] { sourcePath }));
				string aliasIdentity = Convert.ToString(comparableIdentity.Invoke(null, new object[] { aliasPath }));
				if (string.IsNullOrWhiteSpace(sourceIdentity) || !string.Equals(sourceIdentity, aliasIdentity, StringComparison.OrdinalIgnoreCase))
				{
					result.Failures.Add("Standard revision primitive audit: two paths to the same file did not resolve to one file identity.");
				}
				string copyPath = Path.Combine(fixtureRoot, "different-file-same-content.rvt");
				File.Copy(sourcePath, copyPath, true);
				string copyIdentity = Convert.ToString(comparableIdentity.Invoke(null, new object[] { copyPath }));
				if (string.Equals(sourceIdentity, copyIdentity, StringComparison.OrdinalIgnoreCase))
				{
					result.Failures.Add("Standard revision primitive audit: a separate file was treated as the same file identity.");
				}
			}
		}
		catch (Exception ex)
		{
			result.Failures.Add("Standard revision primitive audit failed: " + DescribeException(ex));
		}
		finally
		{
			try
			{
				if (Directory.Exists(fixtureRoot)) Directory.Delete(fixtureRoot, true);
			}
			catch
			{
			}
		}
	}

	private static void RunFileGuardPolicyAudit(string hostAssembly, AuditResult result)
	{
		try
		{
			Assembly assembly = Assembly.LoadFrom(hostAssembly);
			Type policyType = assembly.GetType("FamilyBrowserStandardPolicy", true);
			Type guardType = assembly.GetType("FamilyBrowserFileGuardPolicy", true);
			Type targetType = assembly.GetType("FamilyBrowserFileGuardTarget", true);
			Type contextType = assembly.GetType("FamilyBrowserProjectPolicyContext", true);
			Type securityServiceType = assembly.GetType("FamilyBrowserSecurityPolicyService", true);
			object policy = Activator.CreateInstance(policyType);
			object security = policyType.GetProperty("Security").GetValue(policy, null);
			((IList)security.GetType().GetProperty("AdminUsers").GetValue(security, null)).Add("audit-admin");

			object guard = Activator.CreateInstance(guardType);
			SetProperty(guard, "Enabled", true);
			object target = Activator.CreateInstance(targetType);
			SetProperty(target, "Enabled", true);
			SetProperty(target, "FileName", "CentralModel.rvt");
			SetProperty(target, "CentralPath", @"\\BIM-SERVER\bim\CentralModel.rvt");
			SetProperty(target, "Discipline", "Mechanical");
			SetProperty(target, "BlockFamilyLoadAndEdit", true);
			SetProperty(target, "BlockTypeChanges", true);
			SetProperty(target, "BlockNestedOnlyStandalonePlacement", true);
			((IList)guardType.GetProperty("Targets").GetValue(guard, null)).Add(target);
			SetProperty(policy, "FileGuard", guard);
			string assignedDiscipline = Convert.ToString(targetType.GetProperty("Discipline").GetValue(target, null), CultureInfo.InvariantCulture);
			if (!string.Equals(assignedDiscipline, "Mechanical", StringComparison.Ordinal))
			{
				result.Failures.Add("File Guard audit: the per-file discipline assignment was not retained by the policy target.");
			}
			Type policyStoreType = assembly.GetType("FamilyBrowserStandardPolicyStore", true);
			MethodInfo cloneTargetMethod = policyStoreType.GetMethod("CloneFileGuardTarget", BindingFlags.NonPublic | BindingFlags.Static);
			object clonedTarget = cloneTargetMethod == null ? null : cloneTargetMethod.Invoke(null, new object[] { target });
			string clonedDiscipline = clonedTarget == null
				? string.Empty
				: Convert.ToString(targetType.GetProperty("Discipline").GetValue(clonedTarget, null), CultureInfo.InvariantCulture);
			if (!string.Equals(clonedDiscipline, "Mechanical", StringComparison.Ordinal))
			{
				result.Failures.Add("File Guard audit: cloning the managed policy dropped the per-file discipline assignment.");
			}
			RunFileGuardExcelRoundTripAudit(assembly, policy, guard, result);

			object matchingContext = Activator.CreateInstance(contextType);
			SetProperty(matchingContext, "ProjectTitle", "CentralModel_user");
			SetProperty(matchingContext, "ModelPath", @"C:\Users\audit\Documents\CentralModel_user.rvt");
			SetProperty(matchingContext, "CentralPath", @"\\BIM-SERVER\bim\CentralModel.rvt");
			SetProperty(matchingContext, "IsWorkshared", true);
			object otherContext = Activator.CreateInstance(contextType);
			SetProperty(otherContext, "ProjectTitle", "OtherModel");
			SetProperty(otherContext, "ModelPath", @"C:\Temp\OtherModel.rvt");
			object startupDeferredContext = Activator.CreateInstance(contextType);
			SetProperty(startupDeferredContext, "ProjectTitle", "CentralModel");
			SetProperty(startupDeferredContext, "ModelPath", @"C:\Users\audit\Documents\CentralModel_audit.rvt");
			SetProperty(startupDeferredContext, "CentralPath", "Not checked on startup");
			SetProperty(startupDeferredContext, "IsWorkshared", true);
			object unrelatedSameNameContext = Activator.CreateInstance(contextType);
			SetProperty(unrelatedSameNameContext, "ProjectTitle", "CentralModel");
			SetProperty(unrelatedSameNameContext, "ModelPath", @"C:\Unmanaged\CentralModel.rvt");
			SetProperty(unrelatedSameNameContext, "IsWorkshared", false);

			MethodInfo canNativeGuard = securityServiceType.GetMethod("CanNativeGuard", BindingFlags.Public | BindingFlags.Static);
			if (canNativeGuard == null)
			{
				throw new MissingMethodException("FamilyBrowserSecurityPolicyService.CanNativeGuard was not found.");
			}
			bool editWhenAdminOff = (bool)canNativeGuard.Invoke(null, new[] { policy, "audit-admin", "EditFamilies", matchingContext, (object)false });
			bool loadWhenAdminOff = (bool)canNativeGuard.Invoke(null, new[] { policy, "audit-admin", "LoadFamilies", matchingContext, (object)false });
			bool typeWhenAdminOff = (bool)canNativeGuard.Invoke(null, new[] { policy, "audit-admin", "AddDeleteTypes", matchingContext, (object)false });
			bool editWhenAdminOn = (bool)canNativeGuard.Invoke(null, new[] { policy, "audit-admin", "EditFamilies", matchingContext, (object)true });
			bool loadWhenAdminOn = (bool)canNativeGuard.Invoke(null, new[] { policy, "audit-admin", "LoadFamilies", matchingContext, (object)true });
			bool editOtherFile = (bool)canNativeGuard.Invoke(null, new[] { policy, "audit-admin", "EditFamilies", otherContext, (object)false });
			bool loadBeforeCentralRefresh = (bool)canNativeGuard.Invoke(null, new[] { policy, "audit-admin", "LoadFamilies", startupDeferredContext, (object)false });
			bool editUnrelatedSameName = (bool)canNativeGuard.Invoke(null, new[] { policy, "audit-admin", "EditFamilies", unrelatedSameNameContext, (object)false });
			if (editWhenAdminOff)
			{
				result.Failures.Add("File Guard audit: family load/edit was allowed for a matching central RVT while Admin Mode was off.");
			}
			if (loadWhenAdminOff)
			{
				result.Failures.Add("File Guard audit: the Family Browser load action was allowed for a matching RVT while Admin Mode was off.");
			}
			if (typeWhenAdminOff)
			{
				result.Failures.Add("File Guard audit: type changes were allowed for a matching central RVT while Admin Mode was off.");
			}
			if (!editWhenAdminOn || !loadWhenAdminOn)
			{
				result.Failures.Add("File Guard audit: an administrator was blocked while Admin Mode was on.");
			}
			if (!editOtherFile)
			{
				result.Failures.Add("File Guard audit: a non-target RVT was blocked.");
			}
			if (loadBeforeCentralRefresh)
			{
				result.Failures.Add("File Guard audit: a matching workshared local was allowed before the deferred central-path refresh.");
			}
			if (!editUnrelatedSameName)
			{
				result.Failures.Add("File Guard audit: an unrelated standalone RVT inherited a guard from a same-name target.");
			}

			MethodInfo isTrackingScopeEnabled = securityServiceType.GetMethod("IsProjectElementTrackingScopeEnabled", BindingFlags.Public | BindingFlags.Static);
			MethodInfo isAnyTrackingEnabled = policyStoreType.GetMethod("IsProjectElementChangeTrackingEnabled", BindingFlags.Public | BindingFlags.Static);
			if (isTrackingScopeEnabled == null || isAnyTrackingEnabled == null)
			{
				throw new MissingMethodException("File-specific element tracking policy methods were not found.");
			}
			SetProperty(target, "TrackElementChanges", true);
			bool anyTrackingEnabled = (bool)isAnyTrackingEnabled.Invoke(null, new[] { policy });
			bool matchingTrackingEnabled = (bool)isTrackingScopeEnabled.Invoke(null, new[] { policy, matchingContext });
			bool otherTrackingEnabled = (bool)isTrackingScopeEnabled.Invoke(null, new[] { policy, otherContext });
			if (!anyTrackingEnabled || !matchingTrackingEnabled || otherTrackingEnabled)
			{
				result.Failures.Add("File Guard tracking audit: tracking was not limited to the matching checked RVT.");
			}
			SetProperty(target, "TrackElementChanges", false);
			bool anyTrackingAfterTargetOff = (bool)isAnyTrackingEnabled.Invoke(null, new[] { policy });
			bool matchingTrackingAfterTargetOff = (bool)isTrackingScopeEnabled.Invoke(null, new[] { policy, matchingContext });
			if (anyTrackingAfterTargetOff || matchingTrackingAfterTargetOff)
			{
				result.Failures.Add("File Guard tracking audit: an unchecked RVT remained in element tracking scope.");
			}
			object unregisteredPolicy = Activator.CreateInstance(policyType);
			bool unregisteredAnyTracking = (bool)isAnyTrackingEnabled.Invoke(null, new[] { unregisteredPolicy });
			bool unregisteredScopeTracking = (bool)isTrackingScopeEnabled.Invoke(null, new[] { unregisteredPolicy, matchingContext });
			if (unregisteredAnyTracking || unregisteredScopeTracking)
			{
				result.Failures.Add("File Guard tracking audit: an RVT not registered in Permissions / Guard entered element tracking scope.");
			}
			SetProperty(target, "TrackElementChanges", true);

			SetProperty(target, "BlockFamilyLoadAndEdit", true);
			SetProperty(target, "BlockTypeChanges", false);
			bool familyOnlyEdit = (bool)canNativeGuard.Invoke(null, new[] { policy, "audit-admin", "EditFamilies", matchingContext, (object)false });
			bool familyOnlyType = (bool)canNativeGuard.Invoke(null, new[] { policy, "audit-admin", "AddDeleteTypes", matchingContext, (object)false });
			if (familyOnlyEdit || !familyOnlyType)
			{
				result.Failures.Add("File Guard audit: the family-load/edit flag did not stay independent from the type-change flag.");
			}
			SetProperty(target, "BlockFamilyLoadAndEdit", false);
			SetProperty(target, "BlockTypeChanges", true);
			bool typeOnlyEdit = (bool)canNativeGuard.Invoke(null, new[] { policy, "audit-admin", "EditFamilies", matchingContext, (object)false });
			bool typeOnlyType = (bool)canNativeGuard.Invoke(null, new[] { policy, "audit-admin", "AddDeleteTypes", matchingContext, (object)false });
			if (!typeOnlyEdit || typeOnlyType)
			{
				result.Failures.Add("File Guard audit: the type-change flag did not stay independent from the family-load/edit flag.");
			}
			SetProperty(target, "BlockFamilyLoadAndEdit", true);
			SetProperty(target, "BlockTypeChanges", true);

			Type nativeGuardType = assembly.GetType("FamilyBrowserNativeCommandGuardService", true);
			MethodInfo buildCommandDefinitions = nativeGuardType.GetMethod("BuildCommandDefinitions", BindingFlags.NonPublic | BindingFlags.Static);
			if (buildCommandDefinitions == null)
			{
				throw new MissingMethodException("FamilyBrowserNativeCommandGuardService.BuildCommandDefinitions was not found.");
			}
			object loadDefinition = null;
			object renameDefinition = null;
			foreach (object definition in (IEnumerable)buildCommandDefinitions.Invoke(null, null))
			{
				string key = Convert.ToString(definition.GetType().GetProperty("Key").GetValue(definition, null)) ?? string.Empty;
				if (string.Equals(key, "native-load-family", StringComparison.OrdinalIgnoreCase))
				{
					loadDefinition = definition;
				}
				else if (string.Equals(key, "native-rename-family-or-type", StringComparison.OrdinalIgnoreCase))
				{
					renameDefinition = definition;
				}
			}
			if (loadDefinition == null || !string.Equals(Convert.ToString(loadDefinition.GetType().GetProperty("RequiredPermission").GetValue(loadDefinition, null)), "EditFamilies", StringComparison.OrdinalIgnoreCase))
			{
				result.Failures.Add("File Guard command audit: native Load Family is not routed through EditFamilies.");
			}
			if (renameDefinition == null || !string.Equals(Convert.ToString(renameDefinition.GetType().GetProperty("RequiredPermission").GetValue(renameDefinition, null)), "RenameFamilyOrType", StringComparison.OrdinalIgnoreCase))
			{
				result.Failures.Add("File Guard command audit: family/type rename is not routed through the combined rename permission.");
			}
			if (loadDefinition != null)
			{
				IEnumerable postableNames = (IEnumerable)loadDefinition.GetType().GetProperty("PostableCommandNames").GetValue(loadDefinition, null);
				bool hasLoadFamily = postableNames.Cast<object>().Any(x => string.Equals(Convert.ToString(x), "LoadFamily", StringComparison.OrdinalIgnoreCase));
				if (!hasLoadFamily)
				{
					result.Failures.Add("File Guard command audit: the Revit LoadFamily postable command is not bound.");
				}
			}

			MethodInfo notifyAdminMode = nativeGuardType.GetMethods(BindingFlags.Public | BindingFlags.Static).FirstOrDefault(x => x.Name == "NotifyAdminModeChanged" && x.GetParameters().Length == 3);
			MethodInfo canNativeGuardPermission = nativeGuardType.GetMethod("CanNativeGuardPermission", BindingFlags.NonPublic | BindingFlags.Static);
			if (notifyAdminMode == null || canNativeGuardPermission == null)
			{
				throw new MissingMethodException("Native guard Admin/permission audit seam was not found.");
			}
			notifyAdminMode.Invoke(null, new object[] { false, policy, false });
			bool renameWhenAdminOff = (bool)canNativeGuardPermission.Invoke(null, new object[] { policy, "audit-admin", "RenameFamilyOrType", matchingContext });
			notifyAdminMode.Invoke(null, new object[] { true, policy, false });
			bool renameWhenAdminOn = (bool)canNativeGuardPermission.Invoke(null, new object[] { policy, "audit-admin", "RenameFamilyOrType", matchingContext });
			notifyAdminMode.Invoke(null, new object[] { false, policy, false });
			if (renameWhenAdminOff || !renameWhenAdminOn)
			{
				result.Failures.Add("File Guard command audit: family/type rename did not follow Admin OFF/ON state.");
			}

			MethodInfo shouldRecordProtectedChange = nativeGuardType.GetMethod("ShouldRecordProtectedChange", BindingFlags.NonPublic | BindingFlags.Static);
			if (shouldRecordProtectedChange == null)
			{
				throw new MissingMethodException("Native guard missing-baseline change classifier was not found.");
			}
			bool unseenFirstModification = (bool)shouldRecordProtectedChange.Invoke(null, new object[] { "Modified", false, false });
			bool unchangedKnownModification = (bool)shouldRecordProtectedChange.Invoke(null, new object[] { "Modified", true, true });
			bool changedKnownModification = (bool)shouldRecordProtectedChange.Invoke(null, new object[] { "Modified", true, false });
			bool unseenAddition = (bool)shouldRecordProtectedChange.Invoke(null, new object[] { "Added", false, false });
			if (!unseenFirstModification || unchangedKnownModification || !changedKnownModification || !unseenAddition)
			{
				result.Failures.Add("File Guard updater audit: a first modification missing from the partial Family/Type index did not fail closed.");
			}

			Type dashboardType = assembly.GetType("FamilyBrowserDashboardHtmlForm", true);
			MethodInfo resolveEffectiveAdminMode = dashboardType.GetMethod("ResolveEffectiveAdminMode", BindingFlags.NonPublic | BindingFlags.Static);
			if (resolveEffectiveAdminMode == null)
			{
				throw new MissingMethodException("FamilyBrowserDashboardHtmlForm.ResolveEffectiveAdminMode was not found.");
			}
			bool selectedOffWithCapability = (bool)resolveEffectiveAdminMode.Invoke(null, new object[] { false, true });
			bool selectedOnWithCapability = (bool)resolveEffectiveAdminMode.Invoke(null, new object[] { true, true });
			bool selectedOnWithoutCapability = (bool)resolveEffectiveAdminMode.Invoke(null, new object[] { true, false });
			if (selectedOffWithCapability || !selectedOnWithCapability || selectedOnWithoutCapability)
			{
				result.Failures.Add("Admin state audit: effective Admin mode did not keep the user's ON/OFF selection separate from Admin capability.");
			}

			MethodInfo shouldBlockNestedOnlyPlacement = nativeGuardType.GetMethod("ShouldBlockNestedOnlyStandalonePlacement", BindingFlags.NonPublic | BindingFlags.Static);
			if (shouldBlockNestedOnlyPlacement == null)
			{
				throw new MissingMethodException("Nested-only standalone placement policy seam was not found.");
			}
			bool nestedOnlyWhenAdminOff = (bool)shouldBlockNestedOnlyPlacement.Invoke(null, new object[] { target, false });
			bool nestedOnlyWhenAdminOn = (bool)shouldBlockNestedOnlyPlacement.Invoke(null, new object[] { target, true });
			if (!nestedOnlyWhenAdminOff || nestedOnlyWhenAdminOn)
			{
				result.Failures.Add("File Guard audit: nested-only standalone placement did not follow Admin OFF/ON state.");
			}
			MethodInfo shouldBlockNestedOnlyMatch = nativeGuardType.GetMethod("ShouldBlockNestedOnlyPlacementMatch", BindingFlags.NonPublic | BindingFlags.Static);
			Type nestedOnlyMatchType = assembly.GetType("FamilyBrowserNestedOnlyPlacementMatchResult", true);
			Type nestedOnlyStateType = assembly.GetType("FamilyBrowserNestedOnlyPlacementMatchState", true);
			if (shouldBlockNestedOnlyMatch == null)
			{
				throw new MissingMethodException("Nested-only exact fingerprint decision seam was not found.");
			}
			object nestedOnlyMatch = Activator.CreateInstance(nestedOnlyMatchType);
			SetProperty(nestedOnlyMatch, "State", Enum.Parse(nestedOnlyStateType, "PendingVerification"));
			bool pendingBlocked = (bool)shouldBlockNestedOnlyMatch.Invoke(null, new[] { nestedOnlyMatch });
			SetProperty(nestedOnlyMatch, "State", Enum.Parse(nestedOnlyStateType, "VerificationUnavailable"));
			bool unavailableBlocked = (bool)shouldBlockNestedOnlyMatch.Invoke(null, new[] { nestedOnlyMatch });
			SetProperty(nestedOnlyMatch, "State", Enum.Parse(nestedOnlyStateType, "ExactMatch"));
			bool exactBlocked = (bool)shouldBlockNestedOnlyMatch.Invoke(null, new[] { nestedOnlyMatch });
			if (pendingBlocked || unavailableBlocked || !exactBlocked)
			{
				result.Failures.Add("File Guard audit: nested-only placement was not limited to an exact standard fingerprint match.");
			}
			RunNestedOnlyPlacementCatalogAudit(assembly, result);

			MethodInfo securitySamePath = securityServiceType.GetMethod("SamePath", BindingFlags.NonPublic | BindingFlags.Static);
			Type sharedMatcherType = assembly.GetType("FamilyBrowserFileGuardPathMatcher", true);
			MethodInfo sharedSamePath = sharedMatcherType.GetMethod("PathsReferToSameFile", BindingFlags.Public | BindingFlags.Static);
			MethodInfo resolveFileGuardMatch = sharedMatcherType.GetMethod("Resolve", BindingFlags.Public | BindingFlags.Static);
			if (securitySamePath == null || sharedSamePath == null || resolveFileGuardMatch == null)
			{
				throw new MissingMethodException("The shared File Guard path-matching seams were not found.");
			}
			bool exactPathMatch = (bool)sharedSamePath.Invoke(null, new object[] { @"\\BIM-SERVER\bim\Projects\CentralModel.rvt", @"\\BIM-SERVER\bim\Projects\CentralModel.rvt" });
			bool hostAndIpAliasMatch = (bool)sharedSamePath.Invoke(null, new object[] { @"\\BIM-SERVER\bim\Projects\CentralModel.rvt", @"\\10.20.30.40\bim\Projects\CentralModel.rvt" });
			bool differentShareMatches = (bool)sharedSamePath.Invoke(null, new object[] { @"\\BIM-SERVER\bim\Projects\CentralModel.rvt", @"\\10.20.30.40\other-share\Projects\CentralModel.rvt" });
			bool securityAliasMatch = (bool)securitySamePath.Invoke(null, new object[] { @"\\BIM-SERVER\bim\Projects\CentralModel.rvt", @"\\10.20.30.40\bim\Projects\CentralModel.rvt" });
			if (!exactPathMatch)
			{
				result.Failures.Add("File Guard path audit: an exact normalized RVT path did not match itself.");
			}
			if (hostAndIpAliasMatch || securityAliasMatch)
			{
				result.Failures.Add("File Guard path audit: unresolved hostname/IP strings were treated as the same physical RVT without file identity evidence.");
			}
			if (differentShareMatches)
			{
				result.Failures.Add("File Guard path audit: different UNC shares were treated as the same RVT path.");
			}
			string aliasFixture = Path.Combine(Path.GetTempPath(), "kky-file-guard-alias-" + Guid.NewGuid().ToString("N"));
			try
			{
				string firstFolder = Path.Combine(aliasFixture, "first");
				string secondFolder = Path.Combine(aliasFixture, "second");
				Directory.CreateDirectory(firstFolder);
				Directory.CreateDirectory(secondFolder);
				string firstPath = Path.Combine(firstFolder, "CentralModel.rvt");
				string secondPath = Path.Combine(secondFolder, "CentralModel.rvt");
				File.WriteAllText(firstPath, "physical identity audit", Encoding.UTF8);
				if (!CreateHardLink(secondPath, firstPath, IntPtr.Zero))
				{
					throw new InvalidOperationException("Could not create the File Guard physical-alias fixture.");
				}
				bool physicalAliasMatch = (bool)sharedSamePath.Invoke(null, new object[] { firstPath, secondPath });
				if (!physicalAliasMatch)
				{
					result.Failures.Add("File Guard path audit: two paths backed by the same physical file identity did not match.");
				}
			}
			finally
			{
				try
				{
					if (Directory.Exists(aliasFixture)) Directory.Delete(aliasFixture, true);
				}
				catch
				{
				}
			}

			object duplicateTarget = Activator.CreateInstance(targetType);
			SetProperty(duplicateTarget, "Enabled", true);
			SetProperty(duplicateTarget, "FileName", "CentralModel.rvt");
			SetProperty(duplicateTarget, "CentralPath", @"\\SECOND-SERVER\other\CentralModel.rvt");
			SetProperty(duplicateTarget, "BlockFamilyLoadAndEdit", true);
			SetProperty(duplicateTarget, "BlockTypeChanges", true);
			IList guardTargets = (IList)guardType.GetProperty("Targets").GetValue(guard, null);
			guardTargets.Add(duplicateTarget);
			object ambiguousMatch = resolveFileGuardMatch.Invoke(null, new object[] { guard, startupDeferredContext });
			bool ambiguous = (bool)ambiguousMatch.GetType().GetProperty("Ambiguous").GetValue(ambiguousMatch, null);
			object ambiguousTarget = ambiguousMatch.GetType().GetProperty("Target").GetValue(ambiguousMatch, null);
			bool ambiguousIdentityUncertain = (bool)ambiguousMatch.GetType().GetProperty("IdentityUncertain").GetValue(ambiguousMatch, null);
			string ambiguousMatchKind = Convert.ToString(ambiguousMatch.GetType().GetProperty("MatchKind").GetValue(ambiguousMatch, null), CultureInfo.InvariantCulture);
			bool ambiguousFamilyBlock = ambiguousTarget != null && (bool)targetType.GetProperty("BlockFamilyLoadAndEdit").GetValue(ambiguousTarget, null);
			bool ambiguousTypeBlock = ambiguousTarget != null && (bool)targetType.GetProperty("BlockTypeChanges").GetValue(ambiguousTarget, null);
			if (!ambiguous ||
				!ambiguousIdentityUncertain ||
				ambiguousTarget == null ||
				!ambiguousFamilyBlock ||
				!ambiguousTypeBlock ||
				!string.Equals(ambiguousMatchKind, "ConservativeAmbiguousWorksharedNamePendingIdentity", StringComparison.Ordinal))
			{
				result.Failures.Add("File Guard path audit: ambiguous workshared identity did not preserve the strictest combined guard while identity was unavailable.");
			}
			guardTargets.Remove(duplicateTarget);

			object duplicateExactTarget = Activator.CreateInstance(targetType);
			SetProperty(duplicateExactTarget, "Enabled", true);
			SetProperty(duplicateExactTarget, "FileName", "CentralModel.rvt");
			SetProperty(duplicateExactTarget, "CentralPath", @"\\BIM-SERVER\bim\CentralModel.rvt");
			SetProperty(duplicateExactTarget, "Discipline", "Electrical");
			SetProperty(duplicateExactTarget, "BlockFamilyLoadAndEdit", false);
			SetProperty(duplicateExactTarget, "BlockTypeChanges", true);
			guardTargets.Add(duplicateExactTarget);
			object duplicatePathMatch = resolveFileGuardMatch.Invoke(null, new object[] { guard, matchingContext });
			bool duplicatePathAmbiguous = (bool)duplicatePathMatch.GetType().GetProperty("Ambiguous").GetValue(duplicatePathMatch, null);
			object conservativeTarget = duplicatePathMatch.GetType().GetProperty("Target").GetValue(duplicatePathMatch, null);
			bool conservativeFamilyBlock = conservativeTarget != null && (bool)targetType.GetProperty("BlockFamilyLoadAndEdit").GetValue(conservativeTarget, null);
			bool conservativeTypeBlock = conservativeTarget != null && (bool)targetType.GetProperty("BlockTypeChanges").GetValue(conservativeTarget, null);
			string conservativeDiscipline = conservativeTarget == null ? string.Empty : Convert.ToString(targetType.GetProperty("Discipline").GetValue(conservativeTarget, null), CultureInfo.InvariantCulture);
			if (!duplicatePathAmbiguous || conservativeTarget == null || !conservativeFamilyBlock || !conservativeTypeBlock || !string.IsNullOrWhiteSpace(conservativeDiscipline))
			{
				result.Failures.Add("File Guard path audit: duplicate physical targets did not preserve the most restrictive guard while requiring an explicit trade resolution.");
			}
			guardTargets.Remove(duplicateExactTarget);

			Type settingsStoreType = assembly.GetType("FamilyBrowserUserSettingsStore", true);
			MethodInfo getSettingsRoot = settingsStoreType.GetMethod("GetSettingsRoot", BindingFlags.Public | BindingFlags.Static);
			string settingsRoot = Convert.ToString(getSettingsRoot.Invoke(null, new object[] { false })) ?? string.Empty;
			string expectedLocalRoot = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "KKY", "FamilyBrowser", "Settings");
			if (!string.Equals(Path.GetFullPath(settingsRoot), Path.GetFullPath(expectedLocalRoot), StringComparison.OrdinalIgnoreCase))
			{
				result.Failures.Add("User settings audit: Admin/language settings are not stored in the local Family Browser settings folder.");
			}

			Type projectCacheRecordType = assembly.GetType("ProjectScanCacheRecord", true);
			object projectCacheRecord = Activator.CreateInstance(projectCacheRecordType);
			int projectCacheSchema = Convert.ToInt32(projectCacheRecordType.GetProperty("SchemaVersion").GetValue(projectCacheRecord, null), CultureInfo.InvariantCulture);
			PropertyInfo projectRevisionToken = projectCacheRecordType.GetProperty("ProjectDocumentRevisionToken");
			if (projectCacheSchema < 3 || projectRevisionToken == null)
			{
				result.Failures.Add("Project cache audit: live Revit document revision data is missing from the project scan cache schema.");
			}

			Type automaticStatusType = assembly.GetType("FamilyBrowserAutomaticModelCheckStatus", true);
			object automaticStatus = Activator.CreateInstance(automaticStatusType);
			int automaticStatusSchema = Convert.ToInt32(automaticStatusType.GetProperty("SchemaVersion").GetValue(automaticStatus, null), CultureInfo.InvariantCulture);
			if (automaticStatusSchema < 2 || automaticStatusType.GetProperty("ProgressCurrent") == null || automaticStatusType.GetProperty("ProgressTotal") == null)
			{
				result.Failures.Add("Automatic model check audit: persistent progress fields are missing from the status schema.");
			}
			if (assembly.GetType("FamilyBrowserAutomaticModelCheckProgressWindow", false) == null)
			{
				result.Failures.Add("Automatic model check audit: the synchronous Revit scan has no visible progress surface.");
			}

			Type errorHelpType = assembly.GetType("FamilyBrowserErrorHelp", true);
			MethodInfo buildError = errorHelpType.GetMethod("Build", BindingFlags.Public | BindingFlags.Static);
			object friendlyError = buildError.Invoke(null, new object[]
			{
				"파일별 권한 적용",
				new InvalidOperationException("Family Browser 정책 저장: Family Browser 스캔 데이터는 로컬 C fallback 폴더에 저장하지 않습니다. 홈페이지 경로를 다시 확인해서 공용 관리 폴더를 연결하세요."),
				@"C:\Temp\audit.log",
				true
			});
			string summary = Convert.ToString(friendlyError.GetType().GetProperty("Summary").GetValue(friendlyError, null)) ?? string.Empty;
			if (summary.IndexOf("공용 관리 폴더", StringComparison.Ordinal) < 0)
			{
				result.Failures.Add("Managed-folder error audit: the user-facing error did not identify the unavailable managed folder.");
			}
		}
		catch (Exception ex)
		{
			result.Failures.Add("File Guard policy audit threw " + DescribeException(ex));
		}
	}

	private static void RunFileGuardExcelRoundTripAudit(Assembly assembly, object standardPolicy, object fileGuard, AuditResult result)
	{
		string workbookPath = Path.Combine(Path.GetTempPath(), "kky-family-browser-file-guard-audit-" + Guid.NewGuid().ToString("N") + ".xlsx");
		try
		{
			Type exportServiceType = assembly.GetType("FamilyBrowserPermissionExcelExportService", true);
			Type importServiceType = assembly.GetType("FamilyBrowserFileGuardExcelService", true);
			MethodInfo exportMethod = exportServiceType.GetMethod("ExportFileGuardPolicy", BindingFlags.Public | BindingFlags.Static);
			MethodInfo importMethod = importServiceType.GetMethod("Import", BindingFlags.Public | BindingFlags.Static);
			if (exportMethod == null || importMethod == null)
			{
				throw new MissingMethodException("File Guard Excel export/import methods were not found.");
			}

			exportMethod.Invoke(null, new[] { fileGuard, workbookPath, (object)true });
			object importResult = importMethod.Invoke(null, new[] { workbookPath, standardPolicy, "audit-admin", (object)true });
			int importedCount = Convert.ToInt32(importResult.GetType().GetProperty("ImportedRowCount").GetValue(importResult, null), CultureInfo.InvariantCulture);
			int skippedCount = Convert.ToInt32(importResult.GetType().GetProperty("SkippedRowCount").GetValue(importResult, null), CultureInfo.InvariantCulture);
			object importedPolicy = importResult.GetType().GetProperty("Policy").GetValue(importResult, null);
			IList importedTargets = (IList)importedPolicy.GetType().GetProperty("Targets").GetValue(importedPolicy, null);
			if (importedCount != 1 || skippedCount != 0 || importedTargets == null || importedTargets.Count != 1)
			{
				result.Failures.Add("File Guard Excel audit: the exported policy did not import as exactly one valid RVT row.");
				return;
			}

			object importedTarget = importedTargets[0];
			string discipline = Convert.ToString(importedTarget.GetType().GetProperty("Discipline").GetValue(importedTarget, null), CultureInfo.InvariantCulture);
			bool blockFamily = Convert.ToBoolean(importedTarget.GetType().GetProperty("BlockFamilyLoadAndEdit").GetValue(importedTarget, null), CultureInfo.InvariantCulture);
			bool blockTypes = Convert.ToBoolean(importedTarget.GetType().GetProperty("BlockTypeChanges").GetValue(importedTarget, null), CultureInfo.InvariantCulture);
			bool blockNestedOnly = Convert.ToBoolean(importedTarget.GetType().GetProperty("BlockNestedOnlyStandalonePlacement").GetValue(importedTarget, null), CultureInfo.InvariantCulture);
			if (!string.Equals(discipline, "Mechanical", StringComparison.Ordinal) || !blockFamily || !blockTypes || !blockNestedOnly)
			{
				result.Failures.Add("File Guard Excel audit: discipline or guard flags changed during Korean XLSX export/import round-trip.");
			}

			IList sourceTargets = (IList)fileGuard.GetType().GetProperty("Targets").GetValue(fileGuard, null);
			object sourceTarget = sourceTargets == null || sourceTargets.Count == 0 ? null : sourceTargets[0];
			if (sourceTarget != null)
			{
				SetProperty(sourceTarget, "Discipline", string.Empty);
				try
				{
					exportMethod.Invoke(null, new[] { fileGuard, workbookPath, (object)true });
					object blankTradeResult = importMethod.Invoke(null, new[] { workbookPath, standardPolicy, "audit-admin", (object)true });
					int blankTradeImported = Convert.ToInt32(blankTradeResult.GetType().GetProperty("ImportedRowCount").GetValue(blankTradeResult, null), CultureInfo.InvariantCulture);
					int blankTradeSkipped = Convert.ToInt32(blankTradeResult.GetType().GetProperty("SkippedRowCount").GetValue(blankTradeResult, null), CultureInfo.InvariantCulture);
					if (blankTradeImported != 0 || blankTradeSkipped != 1)
					{
						result.Failures.Add("File Guard Excel audit: a blank per-file trade silently inherited the currently selected standard trade.");
					}
				}
				finally
				{
					SetProperty(sourceTarget, "Discipline", "Mechanical");
				}
			}
		}
		finally
		{
			try
			{
				if (File.Exists(workbookPath)) File.Delete(workbookPath);
			}
			catch
			{
			}
		}
	}

	private static void RunNestedOnlyPlacementCatalogAudit(Assembly assembly, AuditResult result)
	{
		Type snapshotType = assembly.GetType("StandardLibrarySnapshot", true);
		Type familyType = assembly.GetType("StandardLoadableFamilySnapshotItem", true);
		Type nestedType = assembly.GetType("StandardNestedLoadableFamilySnapshotItem", true);
		Type catalogStoreType = assembly.GetType("FamilyBrowserNestedOnlyPlacementCatalogStore", true);
		object snapshot = Activator.CreateInstance(snapshotType);
		SetProperty(snapshot, "SnapshotMode", "Precise");
		IList families = (IList)snapshotType.GetProperty("LoadableFamilies").GetValue(snapshot, null);

		object parent = Activator.CreateInstance(familyType);
		SetProperty(parent, "FamilyName", "ParentAssembly");
		SetProperty(parent, "CategoryName", "Generic Models");
		SetProperty(parent, "CategoryId", "-2000151");
		SetProperty(parent, "StandalonePlacementUsageCaptured", true);
		SetProperty(parent, "StandaloneInstanceCount", 2);
		IList children = (IList)familyType.GetProperty("NestedLoadableFamilies").GetValue(parent, null);
		object nestedOnlyReference = Activator.CreateInstance(nestedType);
		SetProperty(nestedOnlyReference, "FamilyName", "NestedOnlyChild");
		SetProperty(nestedOnlyReference, "CategoryName", "Generic Models");
		SetProperty(nestedOnlyReference, "CategoryId", "-2000151");
		children.Add(nestedOnlyReference);
		object independentlyUsedReference = Activator.CreateInstance(nestedType);
		SetProperty(independentlyUsedReference, "FamilyName", "NestedAndStandalone");
		SetProperty(independentlyUsedReference, "CategoryName", "Generic Models");
		SetProperty(independentlyUsedReference, "CategoryId", "-2000151");
		children.Add(independentlyUsedReference);
		families.Add(parent);

		object nestedOnly = Activator.CreateInstance(familyType);
		SetProperty(nestedOnly, "FamilyName", "NestedOnlyChild");
		SetProperty(nestedOnly, "CategoryName", "Generic Models");
		SetProperty(nestedOnly, "CategoryId", "-2000151");
		SetProperty(nestedOnly, "ContentFingerprint", "AA11BB22");
		SetProperty(nestedOnly, "IsShared", true);
		SetProperty(nestedOnly, "IsNestedLoadableChild", true);
		SetProperty(nestedOnly, "StandalonePlacementUsageCaptured", true);
		SetProperty(nestedOnly, "StandaloneInstanceCount", 0);
		families.Add(nestedOnly);

		object nestedAndStandalone = Activator.CreateInstance(familyType);
		SetProperty(nestedAndStandalone, "FamilyName", "NestedAndStandalone");
		SetProperty(nestedAndStandalone, "CategoryName", "Generic Models");
		SetProperty(nestedAndStandalone, "CategoryId", "-2000151");
		SetProperty(nestedAndStandalone, "ContentFingerprint", "CC33DD44");
		SetProperty(nestedAndStandalone, "IsShared", true);
		SetProperty(nestedAndStandalone, "IsNestedLoadableChild", true);
		SetProperty(nestedAndStandalone, "StandalonePlacementUsageCaptured", true);
		SetProperty(nestedAndStandalone, "StandaloneInstanceCount", 1);
		families.Add(nestedAndStandalone);

		MethodInfo build = catalogStoreType.GetMethod("Build", BindingFlags.Public | BindingFlags.Static);
		MethodInfo contains = catalogStoreType.GetMethod("Contains", BindingFlags.Public | BindingFlags.Static);
		if (build == null || contains == null)
		{
			throw new MissingMethodException("Nested-only placement catalog audit seam is incomplete.");
		}
		object catalog = build.Invoke(null, new object[] { snapshot, @"C:\Temp\standard-snapshot-audit.json" });
		int schemaVersion = Convert.ToInt32(catalog.GetType().GetProperty("SchemaVersion").GetValue(catalog, null));
		bool complete = Convert.ToBoolean(catalog.GetType().GetProperty("IsComplete").GetValue(catalog, null));
		bool nestedOnlyMatched = (bool)contains.Invoke(null, new object[] { catalog, "-2000151", "Generic Models", "NestedOnlyChild" });
		bool standaloneMatched = (bool)contains.Invoke(null, new object[] { catalog, "-2000151", "Generic Models", "NestedAndStandalone" });
		bool wrongCategoryMatched = (bool)contains.Invoke(null, new object[] { catalog, "-2009999", "Doors", "NestedOnlyChild" });
		if (schemaVersion != 2 || !complete || !nestedOnlyMatched || standaloneMatched || wrongCategoryMatched)
		{
			result.Failures.Add("Nested-only placement catalog audit: exclusive nested use, standalone use, or category identity was classified incorrectly.");
		}
		Type fingerprintPolicyType = assembly.GetType("FamilyBrowserNestedOnlyPlacementFingerprintPolicy", true);
		MethodInfo isExactMatch = fingerprintPolicyType.GetMethod("IsExactMatch", BindingFlags.Public | BindingFlags.Static);
		MethodInfo findEntry = catalogStoreType.GetMethod("FindEntry", BindingFlags.Public | BindingFlags.Static);
		object catalogEntry = findEntry.Invoke(null, new object[] { catalog, "-2000151", "Generic Models", "NestedOnlyChild" });
		bool exactFingerprintMatched = (bool)isExactMatch.Invoke(null, new object[] { catalogEntry, "aa11bb22", true });
		bool differentFingerprintMatched = (bool)isExactMatch.Invoke(null, new object[] { catalogEntry, "different", true });
		bool nonSharedProjectMatched = (bool)isExactMatch.Invoke(null, new object[] { catalogEntry, "AA11BB22", false });
		if (!exactFingerprintMatched || differentFingerprintMatched || nonSharedProjectMatched)
		{
			result.Failures.Add("Nested-only placement catalog audit: precise fingerprint or shared-family equality was not enforced.");
		}

		SetProperty(nestedOnly, "StandalonePlacementUsageCaptured", false);
		object legacyCatalog = build.Invoke(null, new object[] { snapshot, @"C:\Temp\legacy-standard-snapshot-audit.json" });
		bool legacyComplete = Convert.ToBoolean(legacyCatalog.GetType().GetProperty("IsComplete").GetValue(legacyCatalog, null));
		bool legacyMatched = (bool)contains.Invoke(null, new object[] { legacyCatalog, "-2000151", "Generic Models", "NestedOnlyChild" });
		if (legacyComplete || legacyMatched)
		{
			result.Failures.Add("Nested-only placement catalog audit: legacy snapshots without standalone-usage metadata were not kept fail-open.");
		}
	}

	private static void RunManagedFolderSetupAudit(string hostAssembly, AuditResult result)
	{
		try
		{
			Assembly assembly = Assembly.LoadFrom(hostAssembly);
			Type setupType = assembly.GetType("FamilyBrowserManagedFolderSetupService", true);
			MethodInfo isNetworkShare = setupType.GetMethod("IsInternalNetworkShare", BindingFlags.Public | BindingFlags.Static);
			MethodInfo pointerPath = setupType.GetMethod("GetPointerPath", BindingFlags.Public | BindingFlags.Static);
			MethodInfo rootFromPolicy = setupType.GetMethod("ResolveRootFromPolicyPath", BindingFlags.Public | BindingFlags.Static);
			MethodInfo analyzeMigration = setupType.GetMethod("AnalyzeMigration", BindingFlags.Public | BindingFlags.Static);
			MethodInfo migrateToHomepage = setupType.GetMethod("MigrateToHomepage", BindingFlags.Public | BindingFlags.Static);
			if (isNetworkShare == null || pointerPath == null || rootFromPolicy == null || analyzeMigration == null || migrateToHomepage == null)
			{
				throw new MissingMethodException("FamilyBrowserManagedFolderSetupService audit seam is incomplete.");
			}
			bool uncAccepted = (bool)isNetworkShare.Invoke(null, new object[] { @"\\audit-server\bim\KKY-FamilyBrowser" });
			bool localRejected = !(bool)isNetworkShare.Invoke(null, new object[] { @"C:\Temp\KKY-FamilyBrowser" });
			if (!uncAccepted || !localRejected)
			{
				result.Failures.Add("Managed folder audit: internal UNC acceptance or local-path rejection failed.");
			}
			string persistedPointer = Convert.ToString(pointerPath.Invoke(null, null)) ?? string.Empty;
			if (!string.Equals(Path.GetFileName(persistedPointer), "managed-folder-override.txt", StringComparison.OrdinalIgnoreCase))
			{
				result.Failures.Add("Managed folder audit: the per-user persisted override pointer path is missing.");
			}
			string resolvedRoot = Convert.ToString(rootFromPolicy.Invoke(null, new object[] { @"\\audit-server\bim\KKY-FamilyBrowser\Config\standard-policy.json" })) ?? string.Empty;
			if (!string.Equals(resolvedRoot.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar), @"\\audit-server\bim\KKY-FamilyBrowser", StringComparison.OrdinalIgnoreCase))
			{
				result.Failures.Add("Managed folder audit: policy path did not resolve back to its management root.");
			}

			string fixtureRoot = Path.Combine(Path.GetTempPath(), "KKY-FamilyBrowser-ManagedFolderMigration-" + Guid.NewGuid().ToString("N"));
			try
			{
				string sourceRoot = Path.Combine(fixtureRoot, "source");
				string destinationRoot = Path.Combine(fixtureRoot, "destination");
				string sourcePolicy = Path.Combine(sourceRoot, "Config", "standard-policy.json");
				string sourceSnapshot = Path.Combine(sourceRoot, "RevitVersions", "Rvt2025", "Snapshots", "audit.json");
				string sourceProjectCatalog = Path.Combine(sourceRoot, "ProjectCatalogs", "project-a", "accepted.json");
				string sourceRevisionManifest = Path.Combine(sourceRoot, "StandardRevisionManifests", "standard-rvt-revision-audit.json");
				string destinationPolicy = Path.Combine(destinationRoot, "Config", "standard-policy.json");
				Directory.CreateDirectory(Path.GetDirectoryName(sourcePolicy));
				Directory.CreateDirectory(Path.GetDirectoryName(sourceSnapshot));
				Directory.CreateDirectory(Path.GetDirectoryName(sourceProjectCatalog));
				Directory.CreateDirectory(Path.GetDirectoryName(sourceRevisionManifest));
				Directory.CreateDirectory(destinationRoot);
				Func<string, string> jsonEscape = value => (value ?? string.Empty).Replace("\\", "\\\\").Replace("\"", "\\\"");
				File.WriteAllText(sourcePolicy, "{\"managedRoot\":\"" + jsonEscape(sourceRoot) + "\"}", new UTF8Encoding(false));
				File.WriteAllText(sourceSnapshot, "{\"preview\":\"" + jsonEscape(Path.Combine(sourceRoot, "RevitVersions", "Rvt2025", "Images", "audit.png")) + "\"}", new UTF8Encoding(false));
				File.WriteAllText(sourceProjectCatalog, "{\"identityPath\":\"" + jsonEscape(Path.Combine(sourceRoot, "RevitVersions", "Rvt2025", "Projects", "project-a.rvt")) + "\"}", new UTF8Encoding(false));
				File.WriteAllText(sourceRevisionManifest, "{\"canonicalPath\":\"" + jsonEscape(Path.Combine(sourceRoot, "RevitVersions", "Rvt2025", "Standards", "audit.rvt")) + "\"}", new UTF8Encoding(false));

				object analysis = analyzeMigration.Invoke(null, new object[] { sourceRoot, destinationPolicy });
				bool canMigrate = Convert.ToBoolean(analysis.GetType().GetProperty("CanMigrate").GetValue(analysis, null));
				int copyCount = Convert.ToInt32(analysis.GetType().GetProperty("CopyFileCount").GetValue(analysis, null));
				int rebaseCount = Convert.ToInt32(analysis.GetType().GetProperty("RebasedJsonFileCount").GetValue(analysis, null));
				if (!canMigrate || copyCount != 4 || rebaseCount != 4)
				{
					result.Failures.Add("Managed folder migration audit: preflight did not include policy, snapshot, project catalog, and Standard RVT revision state with path rebasing.");
				}

				object migration = migrateToHomepage.Invoke(null, new object[] { sourceRoot, destinationPolicy, "ui-audit" });
				bool migrationSuccess = Convert.ToBoolean(migration.GetType().GetProperty("Success").GetValue(migration, null));
				string migrationIssue = Convert.ToString(migration.GetType().GetProperty("Issue").GetValue(migration, null)) ?? string.Empty;
				string migratedPolicyText = File.Exists(destinationPolicy) ? File.ReadAllText(destinationPolicy, Encoding.UTF8) : string.Empty;
				string migratedSnapshotPath = Path.Combine(destinationRoot, "RevitVersions", "Rvt2025", "Snapshots", "audit.json");
				string migratedSnapshotText = File.Exists(migratedSnapshotPath) ? File.ReadAllText(migratedSnapshotPath, Encoding.UTF8) : string.Empty;
				string migratedProjectCatalogPath = Path.Combine(destinationRoot, "ProjectCatalogs", "project-a", "accepted.json");
				string migratedRevisionManifestPath = Path.Combine(destinationRoot, "StandardRevisionManifests", "standard-rvt-revision-audit.json");
				string migratedProjectCatalogText = File.Exists(migratedProjectCatalogPath) ? File.ReadAllText(migratedProjectCatalogPath, Encoding.UTF8) : string.Empty;
				string migratedRevisionManifestText = File.Exists(migratedRevisionManifestPath) ? File.ReadAllText(migratedRevisionManifestPath, Encoding.UTF8) : string.Empty;
				bool sourceRetained = File.Exists(sourcePolicy);
				bool policyRebased = migratedPolicyText.IndexOf(jsonEscape(destinationRoot), StringComparison.OrdinalIgnoreCase) >= 0 && migratedPolicyText.IndexOf(jsonEscape(sourceRoot), StringComparison.OrdinalIgnoreCase) < 0;
				bool snapshotRebased = migratedSnapshotText.IndexOf(jsonEscape(destinationRoot), StringComparison.OrdinalIgnoreCase) >= 0 && migratedSnapshotText.IndexOf(jsonEscape(sourceRoot), StringComparison.OrdinalIgnoreCase) < 0;
				bool projectCatalogRebased = migratedProjectCatalogText.IndexOf(jsonEscape(destinationRoot), StringComparison.OrdinalIgnoreCase) >= 0 && migratedProjectCatalogText.IndexOf(jsonEscape(sourceRoot), StringComparison.OrdinalIgnoreCase) < 0;
				bool revisionManifestRebased = migratedRevisionManifestText.IndexOf(jsonEscape(destinationRoot), StringComparison.OrdinalIgnoreCase) >= 0 && migratedRevisionManifestText.IndexOf(jsonEscape(sourceRoot), StringComparison.OrdinalIgnoreCase) < 0;
				if (!migrationSuccess || !sourceRetained || !policyRebased || !snapshotRebased || !projectCatalogRebased || !revisionManifestRebased)
				{
					result.Failures.Add("Managed folder migration JSON path rebasing audit failed: success=" + migrationSuccess
						+ ", sourceRetained=" + sourceRetained
						+ ", policyRebased=" + policyRebased
						+ ", snapshotRebased=" + snapshotRebased
						+ ", projectCatalogRebased=" + projectCatalogRebased
						+ ", revisionManifestRebased=" + revisionManifestRebased
						+ (string.IsNullOrWhiteSpace(migrationIssue) ? string.Empty : ", issue=" + migrationIssue));
				}

				string conflictSource = Path.Combine(fixtureRoot, "conflict-source");
				string conflictDestination = Path.Combine(fixtureRoot, "conflict-destination");
				string conflictSourcePolicy = Path.Combine(conflictSource, "Config", "standard-policy.json");
				string conflictDestinationPolicy = Path.Combine(conflictDestination, "Config", "standard-policy.json");
				Directory.CreateDirectory(Path.GetDirectoryName(conflictSourcePolicy));
				Directory.CreateDirectory(Path.GetDirectoryName(conflictDestinationPolicy));
				File.WriteAllText(conflictSourcePolicy, "{\"owner\":\"TEST\"}", new UTF8Encoding(false));
				File.WriteAllText(conflictDestinationPolicy, "{\"owner\":\"HOMEPAGE\"}", new UTF8Encoding(false));
				object conflictAnalysis = analyzeMigration.Invoke(null, new object[] { conflictSource, conflictDestinationPolicy });
				bool conflictCanMigrate = Convert.ToBoolean(conflictAnalysis.GetType().GetProperty("CanMigrate").GetValue(conflictAnalysis, null));
				int blockingConflicts = Convert.ToInt32(conflictAnalysis.GetType().GetProperty("BlockingConflictCount").GetValue(conflictAnalysis, null));
				if (conflictCanMigrate || blockingConflicts < 1 || File.ReadAllText(conflictDestinationPolicy, Encoding.UTF8).IndexOf("HOMEPAGE", StringComparison.Ordinal) < 0)
				{
					result.Failures.Add("Managed folder migration audit: a different homepage policy was not protected from overwrite.");
				}

				string rollbackSource = Path.Combine(fixtureRoot, "rollback-source");
				string rollbackDestination = Path.Combine(fixtureRoot, "rollback-destination");
				string rollbackSourceFirst = Path.Combine(rollbackSource, "Config", "a-first.json");
				string rollbackSourceBlocked = Path.Combine(rollbackSource, "RevitVersions", "z-blocked.json");
				string rollbackDestinationPolicy = Path.Combine(rollbackDestination, "Config", "standard-policy.json");
				string rollbackDestinationFirst = Path.Combine(rollbackDestination, "Config", "a-first.json");
				string rollbackDestinationBlocked = Path.Combine(rollbackDestination, "RevitVersions", "z-blocked.json");
				Directory.CreateDirectory(Path.GetDirectoryName(rollbackSourceFirst));
				Directory.CreateDirectory(Path.GetDirectoryName(rollbackSourceBlocked));
				Directory.CreateDirectory(rollbackDestinationBlocked);
				File.WriteAllText(rollbackSourceFirst, "{\"order\":1}", new UTF8Encoding(false));
				File.WriteAllText(rollbackSourceBlocked, "{\"order\":2}", new UTF8Encoding(false));
				object rollbackMigration = migrateToHomepage.Invoke(null, new object[] { rollbackSource, rollbackDestinationPolicy, "ui-audit" });
				bool rollbackSuccess = Convert.ToBoolean(rollbackMigration.GetType().GetProperty("Success").GetValue(rollbackMigration, null));
				int rolledBackFiles = Convert.ToInt32(rollbackMigration.GetType().GetProperty("RolledBackFileCount").GetValue(rollbackMigration, null));
				int rollbackFailedFiles = Convert.ToInt32(rollbackMigration.GetType().GetProperty("RollbackFailedFileCount").GetValue(rollbackMigration, null));
				if (rollbackSuccess || rolledBackFiles < 1 || rollbackFailedFiles != 0 || File.Exists(rollbackDestinationFirst))
				{
					result.Failures.Add("Managed folder migration rollback audit: a partial destination copy survived a deterministic mid-copy failure.");
				}
			}
			finally
			{
				try
				{
					if (Directory.Exists(fixtureRoot))
					{
						Directory.Delete(fixtureRoot, true);
					}
				}
				catch
				{
				}
			}
		}
		catch (Exception ex)
		{
			result.Failures.Add("Managed folder setup audit failed: " + ex.GetBaseException().Message);
		}
	}

	private static void RunFamilyEditDialogGuardAudit(string hostAssembly, AuditResult result)
	{
		string workbookPath = string.Empty;
		try
		{
			Assembly assembly = Assembly.LoadFrom(hostAssembly);
			Type guardType = assembly.GetType("FamilyThumbnailConstraintDialogGuard", true);
			MethodInfo resolver = guardType.GetMethod("ResolveFamilyEditDialogActionForAudit", BindingFlags.Public | BindingFlags.Static);
			if (resolver == null)
			{
				throw new MissingMethodException("FamilyThumbnailConstraintDialogGuard.ResolveFamilyEditDialogActionForAudit was not found.");
			}

			Action<string, string, string[], bool, string> assertAction = (title, body, buttons, activeScope, expected) =>
			{
				string actual = Convert.ToString(resolver.Invoke(null, new object[] { title, body, buttons, activeScope })) ?? string.Empty;
				if (!string.Equals(actual, expected, StringComparison.Ordinal))
				{
					result.Failures.Add("Family edit dialog audit expected " + expected + " but received " + actual + " for: " + body);
				}
			};

			assertAction("Revit", "Opening not cutting anything.", new[] { "OK" }, true, "Confirm|OpeningNotCuttingAnything");
			assertAction("Revit", "Previously unseen family-edit warning text.", new[] { "Confirm" }, true, "Confirm|StandardFamilyScanOkDialog");
			assertAction("Revit", "Delete Instance or Delete Type to continue.", new[] { "Delete Instance", "Delete Type", "Cancel" }, true, "Cancel|DeleteInstanceOrType");
			assertAction("Revit", "Delete Instance or Delete Type to continue.", new[] { "idok:Delete Instance", "Delete Type", "Cancel" }, true, "Cancel|DeleteInstanceOrType");
			assertAction("Revit", "Delete Type is available, but OK can continue safely.", new[] { "Delete Type", "OK", "Cancel" }, true, "Confirm|StandardFamilyScanOkDialog");
			assertAction("Revit", "Unrelated dialog outside family edit.", new[] { "OK" }, false, "None|");
			assertAction("Revit", "Unrelated dialog with only Cancel.", new[] { "Cancel" }, true, "None|");

			Type reportType = assembly.GetType("ProjectStandardComparisonReport", true);
			Type recordType = assembly.GetType("FamilyThumbnailAutoConfirmedDialogRecord", true);
			Type exportType = assembly.GetType("ProjectComparisonReviewExcelExportService", true);
			object report = Activator.CreateInstance(reportType);
			object record = Activator.CreateInstance(recordType);
			SetProperty(record, "ConfirmedAtUtc", "2026-07-15T00:00:00.0000000Z");
			SetProperty(record, "CategoryName", "Doors");
			SetProperty(record, "FamilyName", "AUDIT_OPENING_FAMILY");
			SetProperty(record, "ActionTaken", "OK");
			SetProperty(record, "Reason", "OpeningNotCuttingAnything");
			SetProperty(record, "OverrideResult", "Win32Click:OK");
			SetProperty(record, "AvailableButtons", "enabled:1:OK");
			SetProperty(record, "DialogText", "Opening not cutting anything.");
			Array dialogRecords = Array.CreateInstance(recordType, 1);
			dialogRecords.SetValue(record, 0);
			MethodInfo saveReviewList = exportType.GetMethods(BindingFlags.Public | BindingFlags.Static)
				.FirstOrDefault(method => method.Name == "SaveReviewList" && method.GetParameters().Length == 6);
			if (saveReviewList == null)
			{
				throw new MissingMethodException("ProjectComparisonReviewExcelExportService scan-dialog workbook overload was not found.");
			}
			workbookPath = Path.Combine(Path.GetTempPath(), "KKY-FamilyBrowser-DialogAudit-" + Guid.NewGuid().ToString("N") + ".xlsx");
			saveReviewList.Invoke(null, new object[] { workbookPath, report, "Audit", string.Empty, dialogRecords, false });
			using (FileStream stream = new FileStream(workbookPath, FileMode.Open, FileAccess.Read, FileShare.Read))
			using (ZipArchive archive = new ZipArchive(stream, ZipArchiveMode.Read))
			{
				ZipArchiveEntry workbookEntry = archive.GetEntry("xl/workbook.xml");
				ZipArchiveEntry dialogEntry = archive.GetEntry("xl/worksheets/sheet2.xml");
				if (workbookEntry == null || dialogEntry == null)
				{
					result.Failures.Add("Family edit dialog workbook audit: ScanDialogs worksheet is missing.");
				}
				else
				{
					string workbookXml;
					string dialogXml;
					using (StreamReader reader = new StreamReader(workbookEntry.Open(), Encoding.UTF8))
					{
						workbookXml = reader.ReadToEnd();
					}
					using (StreamReader reader = new StreamReader(dialogEntry.Open(), Encoding.UTF8))
					{
						dialogXml = reader.ReadToEnd();
					}
					if (workbookXml.IndexOf("ScanDialogs", StringComparison.Ordinal) < 0 ||
						dialogXml.IndexOf("AUDIT_OPENING_FAMILY", StringComparison.Ordinal) < 0 ||
						dialogXml.IndexOf("Opening not cutting anything", StringComparison.Ordinal) < 0)
					{
						result.Failures.Add("Family edit dialog workbook audit: family, action, or dialog text was not exported.");
					}
				}
			}
		}
		catch (Exception ex)
		{
			result.Failures.Add("Family edit dialog guard audit failed: " + DescribeException(ex));
		}
		finally
		{
			if (!string.IsNullOrWhiteSpace(workbookPath) && File.Exists(workbookPath))
			{
				try
				{
					File.Delete(workbookPath);
				}
				catch
				{
				}
			}
		}
	}

	private static void RunMeasurementUnitPreferenceAudit(string hostAssembly, AuditResult result)
	{
		string auditFolder = string.Empty;
		try
		{
			Assembly assembly = Assembly.LoadFrom(hostAssembly);
			Type preferenceType = assembly.GetType("FamilyBrowserMeasurementUnitPreferenceService", true);
			MethodInfo load = preferenceType.GetMethod("LoadFromPathForAudit", BindingFlags.NonPublic | BindingFlags.Static);
			MethodInfo save = preferenceType.GetMethod("SaveToPathForAudit", BindingFlags.NonPublic | BindingFlags.Static);
			if (load == null || save == null)
			{
				throw new MissingMethodException("FamilyBrowserMeasurementUnitPreferenceService audit seams were not found.");
			}

			auditFolder = Path.Combine(Path.GetTempPath(), "KKY-FamilyBrowser-MeasurementUnitAudit-" + Guid.NewGuid().ToString("N"));
			string preferencePath = Path.Combine(auditFolder, "measurement-unit.txt");
			string initial = Convert.ToString(load.Invoke(null, new object[] { preferencePath })) ?? string.Empty;
			if (!string.Equals(initial, "mm", StringComparison.Ordinal))
			{
				result.Failures.Add("Measurement unit preference did not default to mm.");
			}

			bool inchSaved = (bool)save.Invoke(null, new object[] { preferencePath, "in" });
			string restoredInch = Convert.ToString(load.Invoke(null, new object[] { preferencePath })) ?? string.Empty;
			if (!inchSaved || !string.Equals(restoredInch, "in", StringComparison.Ordinal))
			{
				result.Failures.Add("Measurement unit preference did not restore inch.");
			}

			bool invalidSaved = (bool)save.Invoke(null, new object[] { preferencePath, "cm" });
			string normalizedInvalid = Convert.ToString(load.Invoke(null, new object[] { preferencePath })) ?? string.Empty;
			if (!invalidSaved || !string.Equals(normalizedInvalid, "mm", StringComparison.Ordinal))
			{
				result.Failures.Add("Measurement unit preference did not normalize an unsupported unit to mm.");
			}
		}
		catch (Exception ex)
		{
			result.Failures.Add("Measurement unit preference audit failed: " + DescribeException(ex));
		}
		finally
		{
			if (!string.IsNullOrWhiteSpace(auditFolder) && Directory.Exists(auditFolder))
			{
				try
				{
					Directory.Delete(auditFolder, true);
				}
				catch
				{
				}
			}
		}
	}

	private static void RunProductUpdatePrimitiveAudit(string hostAssembly, AuditResult result)
	{
		string auditFolder = string.Empty;
		try
		{
			Assembly assembly = Assembly.LoadFrom(hostAssembly);
			Type serviceType = assembly.GetType("FamilyBrowserProductUpdateService", true);
			MethodInfo trustUrl = serviceType.GetMethod("IsTrustedInstallerUrlForAudit", BindingFlags.NonPublic | BindingFlags.Static);
			MethodInfo validateFile = serviceType.GetMethod("ValidateInstallerFileForAudit", BindingFlags.NonPublic | BindingFlags.Static);
			if (trustUrl == null || validateFile == null)
			{
				throw new MissingMethodException("FamilyBrowserProductUpdateService audit seams were not found.");
			}

			Func<string, bool> isTrusted = address => (bool)trustUrl.Invoke(null, new object[] { address });
			string trustedInstaller = "https://update.zerokky.com/Release/family-browser/official/KKY_FamilyBrowser_v1.1_Setup.exe";
			if (!isTrusted(trustedInstaller))
			{
				result.Failures.Add("Product update audit rejected the trusted HTTPS installer URL.");
			}
			if (isTrusted("http://update.zerokky.com/Release/family-browser/official/KKY_FamilyBrowser_v1.1_Setup.exe") ||
				isTrusted("https://evil.example/KKY_FamilyBrowser_v1.1_Setup.exe") ||
				isTrusted("https://update.zerokky.com/Release/family-browser/official/latest.json"))
			{
				result.Failures.Add("Product update audit accepted an untrusted installer URL.");
			}

			auditFolder = Path.Combine(Path.GetTempPath(), "KKY-FamilyBrowser-ProductUpdateAudit-" + Guid.NewGuid().ToString("N"));
			Directory.CreateDirectory(auditFolder);
			byte[] installerBytes = new byte[131072];
			new Random(20260721).NextBytes(installerBytes);
			installerBytes[0] = 0x4D;
			installerBytes[1] = 0x5A;
			string installerPath = Path.Combine(auditFolder, "valid-update.exe");
			File.WriteAllBytes(installerPath, installerBytes);

			string validHash;
			using (SHA256 sha = SHA256.Create())
			{
				validHash = BitConverter.ToString(sha.ComputeHash(installerBytes)).Replace("-", string.Empty);
			}
			if (!(bool)validateFile.Invoke(null, new object[] { installerPath, validHash }))
			{
				result.Failures.Add("Product update audit rejected an intact MZ installer with the expected SHA-256.");
			}

			string wrongHash = (validHash[0] == 'A' ? "B" : "A") + validHash.Substring(1);
			if ((bool)validateFile.Invoke(null, new object[] { installerPath, wrongHash }))
			{
				result.Failures.Add("Product update audit accepted an installer with the wrong SHA-256.");
			}

			installerBytes[0] = 0x00;
			string invalidPath = Path.Combine(auditFolder, "invalid-update.exe");
			File.WriteAllBytes(invalidPath, installerBytes);
			string invalidHash;
			using (SHA256 sha = SHA256.Create())
			{
				invalidHash = BitConverter.ToString(sha.ComputeHash(installerBytes)).Replace("-", string.Empty);
			}
			if ((bool)validateFile.Invoke(null, new object[] { invalidPath, invalidHash }))
			{
				result.Failures.Add("Product update audit accepted a non-MZ file.");
			}
		}
		catch (Exception ex)
		{
			result.Failures.Add("Product update primitive audit failed: " + DescribeException(ex));
		}
		finally
		{
			if (!string.IsNullOrWhiteSpace(auditFolder) && Directory.Exists(auditFolder))
			{
				try
				{
					Directory.Delete(auditFolder, true);
				}
				catch
				{
				}
			}
		}
	}

	private static void SetProperty(object target, string propertyName, object value)
	{
		PropertyInfo property = target.GetType().GetProperty(propertyName, BindingFlags.Public | BindingFlags.Instance);
		if (property == null)
		{
			throw new MissingMemberException(target.GetType().FullName, propertyName);
		}
		property.SetValue(target, value, null);
	}

	private static Dictionary<string, string> ParseArgs(string[] args)
	{
		Dictionary<string, string> options = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
		for (int i = 0; i < args.Length; i++)
		{
			string key = args[i] ?? string.Empty;
			if (!key.StartsWith("--", StringComparison.Ordinal))
			{
				continue;
			}
			key = key.Substring(2);
			string value = "true";
			if (i + 1 < args.Length && !args[i + 1].StartsWith("--", StringComparison.Ordinal))
			{
				value = args[++i];
			}
			options[key] = value;
		}
		return options;
	}

	private static string GetOption(Dictionary<string, string> options, string key, string defaultValue)
	{
		string value;
		if (options.TryGetValue(key, out value) && value != null)
		{
			return value;
		}
		return defaultValue;
	}

	private static bool GetBool(Dictionary<string, string> options, string key, bool defaultValue)
	{
		string value = GetOption(options, key, defaultValue ? "true" : "false");
		bool parsed;
		if (bool.TryParse(value, out parsed))
		{
			return parsed;
		}
		return defaultValue;
	}

	private static List<string> BuildDependencyDirs(Dictionary<string, string> options, string hostAssembly)
	{
		List<string> dirs = new List<string>();
		dirs.Add(Path.GetDirectoryName(Path.GetFullPath(hostAssembly)));
		string dependencyArg = GetOption(options, "dependencyDir", string.Empty);
		foreach (string part in dependencyArg.Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries))
		{
			string dir = part.Trim();
			if (dir.Length > 0)
			{
				dirs.Add(dir);
			}
		}
		return dirs.Where(Directory.Exists).Distinct(StringComparer.OrdinalIgnoreCase).ToList();
	}

	private static void ExtendProcessPath(List<string> dependencyDirs)
	{
		string current = Environment.GetEnvironmentVariable("PATH") ?? string.Empty;
		List<string> parts = new List<string>(dependencyDirs);
		parts.Add(current);
		Environment.SetEnvironmentVariable("PATH", string.Join(Path.PathSeparator.ToString(), parts.Where(x => !string.IsNullOrWhiteSpace(x)).ToArray()));
		AddNativeDllSearchDirectories(dependencyDirs);
	}

	private static void AddNativeDllSearchDirectories(List<string> dependencyDirs)
	{
		try
		{
			SetDefaultDllDirectories(LoadLibrarySearchDefaultDirs | LoadLibrarySearchUserDirs);
			foreach (string dir in dependencyDirs)
			{
				if (!string.IsNullOrWhiteSpace(dir) && Directory.Exists(dir))
				{
					AddDllDirectory(dir);
				}
			}
		}
		catch
		{
		}
	}

	private static void RegisterAssemblyResolver(List<string> dependencyDirs)
	{
#if NETFRAMEWORK
		AppDomain.CurrentDomain.AssemblyResolve += delegate(object sender, ResolveEventArgs e)
		{
			return ResolveAssembly(e.Name, dependencyDirs);
		};
#else
		System.Runtime.Loader.AssemblyLoadContext.Default.Resolving += delegate(System.Runtime.Loader.AssemblyLoadContext context, AssemblyName name)
		{
			string path = ResolveAssemblyPath(name.Name, dependencyDirs);
			return string.IsNullOrWhiteSpace(path) ? null : context.LoadFromAssemblyPath(path);
		};
#endif
	}

	private static void PreloadKnownDependencies(List<string> dependencyDirs, AuditResult result)
	{
		foreach (string simpleName in new[] { "RevitAPI", "RevitAPIUI" })
		{
			string path = ResolveAssemblyPath(simpleName, dependencyDirs);
			if (string.IsNullOrWhiteSpace(path))
			{
				result.Warnings.Add("Dependency not found for preload: " + simpleName);
				continue;
			}
			try
			{
				Assembly.LoadFrom(path);
			}
			catch (Exception ex)
			{
				result.Warnings.Add("Dependency preload failed for " + simpleName + ": " + DescribeException(ex));
			}
		}
	}

	private static Assembly ResolveAssembly(string assemblyName, List<string> dependencyDirs)
	{
		AssemblyName name = new AssemblyName(assemblyName);
		string path = ResolveAssemblyPath(name.Name, dependencyDirs);
		return string.IsNullOrWhiteSpace(path) ? null : Assembly.LoadFrom(path);
	}

	private static string ResolveAssemblyPath(string simpleName, List<string> dependencyDirs)
	{
		foreach (string dir in dependencyDirs)
		{
			string candidate = Path.Combine(dir, simpleName + ".dll");
			if (File.Exists(candidate))
			{
				return candidate;
			}
		}
		return string.Empty;
	}

	private static string RenderScenarioHtml(Dictionary<string, string> options, string hostAssembly)
	{
		Assembly assembly = Assembly.LoadFrom(hostAssembly);
		Type scenarioType = assembly.GetType("FamilyBrowserDashboardAuditScenario", true);
		object scenario = Activator.CreateInstance(scenarioType);
		SetScenarioProperty(scenarioType, scenario, "Name", GetOption(options, "scenario", "audit-default"));
		SetScenarioProperty(scenarioType, scenario, "WorkspaceRoot", GetOption(options, "workspaceRoot", string.Empty));
		SetScenarioProperty(scenarioType, scenario, "LanguageCode", GetOption(options, "languageCode", "ko"));
		SetScenarioProperty(scenarioType, scenario, "InitialLanguageCode", GetOption(options, "initialLanguageCode", string.Empty));
		SetScenarioProperty(scenarioType, scenario, "ThemeCode", GetOption(options, "themeCode", "light"));
		SetScenarioProperty(scenarioType, scenario, "ActiveTab", GetOption(options, "activeTab", "home"));
		SetScenarioProperty(scenarioType, scenario, "BrowseDisciplineKey", GetOption(options, "browseDisciplineKey", "Mechanical"));
		SetScenarioProperty(scenarioType, scenario, "PolicyActiveDisciplineKey", GetOption(options, "policyActiveDisciplineKey", string.Empty));
		SetScenarioProperty(scenarioType, scenario, "AdminMode", GetBool(options, "adminMode", true));
		SetScenarioProperty(scenarioType, scenario, "AdminProfile", GetBool(options, "adminProfile", false));
		SetScenarioProperty(scenarioType, scenario, "FileGuardProtected", GetBool(options, "fileGuardProtected", false));
		SetScenarioProperty(scenarioType, scenario, "StandardRvtRegistered", GetBool(options, "standardRvtRegistered", true));
		SetScenarioProperty(scenarioType, scenario, "StandardListRegistered", GetBool(options, "standardListRegistered", true));
		SetScenarioProperty(scenarioType, scenario, "StandardRvtChanged", GetBool(options, "standardRvtChanged", false));
		SetScenarioProperty(scenarioType, scenario, "StandardRvtUnavailable", GetBool(options, "standardRvtUnavailable", false));
		SetScenarioProperty(scenarioType, scenario, "ProjectCatalogBaselineMissing", GetBool(options, "projectCatalogBaselineMissing", false));
		SetScenarioProperty(scenarioType, scenario, "ProjectCatalogChanged", GetBool(options, "projectCatalogChanged", false));
		SetScenarioProperty(scenarioType, scenario, "ProjectCatalogUntracked", GetBool(options, "projectCatalogUntracked", false));
		SetScenarioProperty(scenarioType, scenario, "TrackingPendingCount", ParseInt(GetOption(options, "trackingPendingCount", "0"), 0));
		SetScenarioProperty(scenarioType, scenario, "IncludeRows", GetBool(options, "includeRows", true));
		SetScenarioProperty(scenarioType, scenario, "IncludePendingRows", GetBool(options, "includePendingRows", false));
		SetScenarioProperty(scenarioType, scenario, "IncludeRequests", GetBool(options, "includeRequests", true));
		SetScenarioProperty(scenarioType, scenario, "IncludeUnregistered", GetBool(options, "includeUnregistered", true));
		SetScenarioProperty(scenarioType, scenario, "IncludeReadinessWarning", GetBool(options, "includeReadinessWarning", false));
		SetScenarioProperty(scenarioType, scenario, "ManagedFolderUnavailable", GetBool(options, "managedFolderUnavailable", false));
		SetScenarioProperty(scenarioType, scenario, "ManagedFolderTestOverride", GetBool(options, "managedFolderTestOverride", false));
		SetScenarioProperty(scenarioType, scenario, "HomepageManagedFolderAvailable", GetBool(options, "homepageManagedFolderAvailable", false));
		SetScenarioProperty(scenarioType, scenario, "CompareDetailedSystemTypeComponents", GetBool(options, "compareDetailedSystemTypeComponents", true));
		SetScenarioProperty(scenarioType, scenario, "SyntheticFamilyCount", ParseInt(GetOption(options, "syntheticFamilyCount", "0"), 0));
		SetScenarioProperty(scenarioType, scenario, "SyntheticSystemCount", ParseInt(GetOption(options, "syntheticSystemCount", "0"), 0));
		SetScenarioProperty(scenarioType, scenario, "UserIdentity", GetOption(options, "userIdentity", GetBool(options, "adminMode", true) ? "KKY_UI_AUDIT_ADMIN" : "KKY_UI_AUDIT_MODELER"));
		if (options.ContainsKey("projectPath"))
		{
			SetScenarioProperty(scenarioType, scenario, "ProjectPath", GetOption(options, "projectPath", string.Empty));
		}
		if (options.ContainsKey("centralPath"))
		{
			SetScenarioProperty(scenarioType, scenario, "CentralPath", GetOption(options, "centralPath", string.Empty));
		}

		Type formType = assembly.GetType("FamilyBrowserDashboardHtmlForm", true);
		MethodInfo render = formType.GetMethod("BuildDashboardHtmlForAudit", BindingFlags.Public | BindingFlags.Static);
		if (render == null)
		{
			throw new MissingMethodException("FamilyBrowserDashboardHtmlForm", "BuildDashboardHtmlForAudit");
		}
		return Convert.ToString(render.Invoke(null, new[] { scenario }));
	}

	private static string RenderMessageDialogHtml(Dictionary<string, string> options, string hostAssembly)
	{
		Assembly assembly = Assembly.LoadFrom(hostAssembly);
		Type dialogType = assembly.GetType("FamilyBrowserModernMessageDialog", true);
		MethodInfo render = dialogType.GetMethod("BuildHtmlForThemeAudit", BindingFlags.Public | BindingFlags.Static);
		if (render == null)
		{
			throw new MissingMethodException("FamilyBrowserModernMessageDialog", "BuildHtmlForThemeAudit");
		}
		bool isKorean = !string.Equals(GetOption(options, "languageCode", "ko"), "en", StringComparison.OrdinalIgnoreCase);
		bool automaticResult = string.Equals(GetOption(options, "messageFixture", string.Empty), "auto-result", StringComparison.OrdinalIgnoreCase);
		string message = automaticResult ? BuildAutomaticResultAuditFixture(isKorean) : BuildMessageAuditFixture(isKorean);
		string caption = automaticResult ? (isKorean ? "현재 모델 검사" : "Current Model Check") : (isKorean ? "공종 설정" : "Discipline settings");
		return Convert.ToString(render.Invoke(null, new object[] { isKorean, message, caption, automaticResult ? MessageBoxIcon.Asterisk : MessageBoxIcon.Hand, GetOption(options, "themeCode", "light") }));
	}

	private static string BuildAutomaticResultAuditFixture(bool isKorean)
	{
		if (isKorean)
		{
			return "현재 모델 검사가 완료되었습니다.\r\n\r\n프로젝트: AUDIT_PROJECT\r\n표준: AUDIT_STANDARD\r\n\r\n비교 결과\r\n- 기준 일치: 24\r\n- 조치 필요: 3\r\n- 실패: 1\r\n리포트 저장: C:\\KKY\\Reports\\audit-result.xlsx\r\n검토가 필요한 항목을 확인한 뒤 Excel 결과를 공유하세요.";
		}
		return "Current model check completed.\r\n\r\nProject: AUDIT_PROJECT\r\nStandard: AUDIT_STANDARD\r\n\r\nComparison result\r\n- Matches standard: 24\r\n- Action needed: 3\r\n- Failed: 1\r\nReport saved: C:\\KKY\\Reports\\audit-result.xlsx\r\nReview the action-needed items, then share the Excel result.";
	}

	private static string BuildMessageAuditFixture(bool isKorean)
	{
		if (isKorean)
		{
			return "공종 설정 작업을 완료하지 못했습니다.\r\n\r\n실패 이유\r\n현재 작업 중 처리하지 못한 예외가 발생했습니다. 자세한 원인은 로그 파일에 저장되었습니다.\r\n\r\n지금 할 일\r\n입력값과 관리 경로 연결 상태를 확인한 뒤 다시 실행하세요.\r\n\r\n관리자에게 전달할 정보\r\n같은 문제가 반복되면 아래 정보를 관리자에게 전달하세요.\r\n로그: C:\\KKY\\Logs\\family-browser.log\r\n지원 코드: KKY-FB-AUDIT-001\r\n\r\n기술 정보\r\nSystem.InvalidOperationException: Audit structured message\r\n   at KKY.FamilyBrowser.Audit.Run()";
		}
		return "Discipline settings could not be completed.\r\n\r\nWhy It Failed\r\nAn exception could not be handled during the current operation. The detailed cause was written to the log.\r\n\r\nWhat To Do Now\r\nCheck the input values and managed-path connection, then run the action again.\r\n\r\nSend This To The Administrator\r\nIf the problem repeats, send the following information to the administrator.\r\nLog: C:\\KKY\\Logs\\family-browser.log\r\nSupport code: KKY-FB-AUDIT-001\r\n\r\nTechnical Detail\r\nSystem.InvalidOperationException: Audit structured message\r\n   at KKY.FamilyBrowser.Audit.Run()";
	}

	private static string RenderStartupShellHtml(Dictionary<string, string> options, string hostAssembly)
	{
		Assembly assembly = Assembly.LoadFrom(hostAssembly);
		Type formType = assembly.GetType("FamilyBrowserDashboardHtmlForm", true);
		MethodInfo render = formType.GetMethod("BuildStartupShellHtmlForAudit", BindingFlags.Public | BindingFlags.Static);
		if (render == null)
		{
			throw new MissingMethodException("FamilyBrowserDashboardHtmlForm", "BuildStartupShellHtmlForAudit");
		}
		return Convert.ToString(render.Invoke(null, new object[] { GetOption(options, "languageCode", "ko") }));
	}

	private static void SetScenarioProperty(Type scenarioType, object scenario, string propertyName, object value)
	{
		PropertyInfo property = scenarioType.GetProperty(propertyName, BindingFlags.Public | BindingFlags.Instance);
		if (property == null)
		{
			throw new MissingMemberException(scenarioType.FullName, propertyName);
		}
		property.SetValue(scenario, value, null);
	}

	private static void RunWebBrowserAudit(string html, Dictionary<string, string> options, AuditResult result)
	{
		int width = ParseInt(GetOption(options, "width", "1740"), 1740);
		int height = ParseInt(GetOption(options, "height", "980"), 980);
		int maxClicks = ParseInt(GetOption(options, "maxClicks", "240"), 240);
		using (Form host = new Form())
		using (WebBrowser browser = new WebBrowser())
		{
			host.Text = "KKY Family Browser UI Audit";
			host.ShowInTaskbar = false;
			host.StartPosition = FormStartPosition.Manual;
			host.Left = -32000;
			host.Top = -32000;
			host.Width = width;
			host.Height = height;
			browser.Dock = DockStyle.Fill;
			browser.ScriptErrorsSuppressed = true;
			browser.AllowWebBrowserDrop = false;
			browser.IsWebBrowserContextMenuEnabled = false;
			browser.WebBrowserShortcutsEnabled = false;
			browser.Navigating += delegate(object sender, WebBrowserNavigatingEventArgs e)
			{
				string raw = e.Url == null ? string.Empty : e.Url.ToString();
				string action = NormalizeHostAction(raw);
				if (action.Length > 0)
				{
					result.HostActions.Add(action);
					if (action.StartsWith("preview-inline/", StringComparison.OrdinalIgnoreCase))
					{
						string previewAction = action;
						host.BeginInvoke(new Action(delegate
						{
							ApplyInlinePreviewAction(browser, previewAction, result);
						}));
					}
					e.Cancel = true;
				}
			};
			host.Controls.Add(browser);
			host.Show();
			Stopwatch documentLoadSw = Stopwatch.StartNew();
			LoadHtmlIntoBrowser(browser, html, string.Empty, false);
			documentLoadSw.Stop();
			result.DocumentLoadMilliseconds = documentLoadSw.ElapsedMilliseconds;
			if (browser.Document == null || browser.Document.Body == null)
			{
				result.Failures.Add("WebBrowser did not finish loading the dashboard HTML.");
				return;
			}
			InjectAuditHooks(browser, result);
			Stopwatch dashboardReadySw = Stopwatch.StartNew();
			bool dashboardReady = WaitForDashboardReady(browser, 6000);
			if (!dashboardReady)
			{
				EvaluateInlineScripts(browser, html, result);
				InjectAuditHooks(browser, result);
				ForceDashboardOnLoad(browser, result);
				dashboardReady = WaitForDashboardReady(browser, 3000);
			}
			if (!dashboardReady && DashboardReadyDiagnostic(browser).IndexOf("openAdvancedFilter=function", StringComparison.OrdinalIgnoreCase) < 0)
			{
				string retryPath = GetOption(options, "htmlOut", string.Empty);
				if (string.IsNullOrWhiteSpace(retryPath))
				{
					retryPath = Path.Combine(Path.GetTempPath(), "kky-family-browser-ui-audit-" + Guid.NewGuid().ToString("N") + ".html");
					File.WriteAllText(retryPath, html ?? string.Empty, new UTF8Encoding(false));
				}
				result.Warnings.Add("DocumentText dashboard JS was not ready; retrying by navigating to the rendered HTML file. Diagnostic: " + DashboardReadyDiagnostic(browser));
				LoadHtmlIntoBrowser(browser, html, retryPath, true);
				if (browser.Document == null || browser.Document.Body == null)
				{
					result.Failures.Add("WebBrowser did not finish loading the dashboard HTML after file-navigation retry.");
					return;
				}
				InjectAuditHooks(browser, result);
				dashboardReady = WaitForDashboardReady(browser, 6000);
				if (!dashboardReady)
				{
					EvaluateInlineScripts(browser, html, result);
					InjectAuditHooks(browser, result);
					ForceDashboardOnLoad(browser, result);
					dashboardReady = WaitForDashboardReady(browser, 3000);
				}
			}
			if (!dashboardReady)
			{
				result.Failures.Add("Dashboard JS did not become ready. Diagnostic: " + DashboardReadyDiagnostic(browser));
			}
			dashboardReadySw.Stop();
			result.DashboardReadyMilliseconds = dashboardReadySw.ElapsedMilliseconds;
			if (GetBool(options, "performanceMode", false))
			{
				CheckPerformance(browser, options, result);
			}

			CheckLayout(browser, result);
			CheckOverflowTitleBehavior(browser, result);
			CheckAuditTargetScrollPersistence(browser, result);
			CheckDebugDock(browser, result);
			CheckProjectSubtitle(browser, options, result);
			CheckLanguagePurity(browser, options, result);
			CheckStandardSetupEmptyState(browser, options, result);
			CheckAdminOffFileGuardUi(browser, options, result);
			CheckStandardRevisionState(browser, options, result);
			CheckManagedFolderRecovery(browser, options, result);
			CheckManagedFolderTransition(browser, options, result);
			CheckProjectCatalogState(browser, options, result);
			CheckPendingTrackingQueue(browser, options, result);
			CheckPendingCommitState(browser, options, result);
			CaptureDashboardVisualHtml(browser, options, result);
			CaptureBrowserScreenshot(browser, options, result);
			if (!GetBool(options, "performanceMode", false))
			{
				CheckAutoDetachedDetailAction(options, result);
				CheckNestedFamilyDifferenceState(browser, options, result);
				CheckBrowserDetailContent(browser, options, result);
				CheckBrowserSystemDetailContent(browser, options, result);
				CaptureDetailScreenshot(browser, options, result);
				CheckSearchFilteringKeepsFocus(browser, options, result);
				CheckDetailedFilterResetAcrossBrowserTabs(browser, options, result);
			}
			CheckThemeBehavior(browser, options, result);
			List<Clickable> clickables = CollectClickables(browser);
			result.ClickableCount = clickables.Count;
			int clicked = 0;
			foreach (Clickable clickable in clickables)
			{
				if (clicked >= maxClicks)
				{
					result.Warnings.Add("Click limit reached: " + maxClicks);
					break;
				}
				if (IsHostActionHref(clickable.Href))
				{
					result.HostActionCandidateCount++;
					continue;
				}
				if (!dashboardReady)
				{
					continue;
				}
				if (string.Equals(clickable.Href, "#", StringComparison.Ordinal) || clickable.Href.EndsWith("#", StringComparison.Ordinal) || !string.IsNullOrWhiteSpace(clickable.OnClick))
				{
					clicked++;
					result.BrowserClickCount++;
					string beforeError = BodyAttribute(browser, "data-kkyfb-js-error");
					string beforeAlert = BodyAttribute(browser, "data-kkyfb-last-alert");
					try
					{
						clickable.Element.InvokeMember("click");
						DoEventsFor(180);
					}
					catch (Exception ex)
					{
						result.Failures.Add("Click threw " + ex.GetType().Name + " on " + Describe(clickable) + ": " + ex.Message);
						continue;
					}
					string afterError = BodyAttribute(browser, "data-kkyfb-js-error");
					if (!string.IsNullOrWhiteSpace(afterError) && !string.Equals(beforeError, afterError, StringComparison.Ordinal))
					{
						result.Failures.Add("JS error after click " + Describe(clickable) + ": " + afterError);
					}
					string afterAlert = BodyAttribute(browser, "data-kkyfb-last-alert");
					if (!string.IsNullOrWhiteSpace(afterAlert) && !string.Equals(beforeAlert, afterAlert, StringComparison.Ordinal))
					{
						result.Warnings.Add("Expected/handled alert after click " + Describe(clickable) + ": " + afterAlert);
					}
				}
				else
				{
					result.Failures.Add("Visible clickable has no host href or onclick: " + Describe(clickable));
				}
			}
			if (!GetBool(options, "performanceMode", false))
			{
				CheckBrowseTreeSynchronization(browser, result);
				CheckBrowseRowDisciplineSynchronization(browser, result);
				CheckBrowseDisciplineSwitchFeedback(browser, result);
			}
			browser.Stop();
			browser.DocumentText = "<html><body></body></html>";
			host.Close();
		}
	}

	private static void CheckBrowseTreeSynchronization(WebBrowser browser, AuditResult result)
	{
		string validation = EvalString(browser, @"
(function(){
  var tab=window.currentTab||'';
  if(tab!='families'&&tab!='systems')return 'SKIP';
  function directText(e){var value='',nodes=e?e.childNodes:[];for(var i=0;i<nodes.length;i++)if(nodes[i].nodeType==3)value+=nodes[i].nodeValue||'';return value.replace(/^\s+|\s+$/g,'');}
  var chips=document.getElementsByName('disciplineFilter');
  var selected='',selectedLabel='';
  for(var i=0;i<chips.length;i++)if((' '+(chips[i].className||'')+' ').indexOf(' active ')>=0){selected=chips[i].getAttribute('data-discipline')||'';selectedLabel=directText(chips[i]);}
  if(!selected)return 'SKIP';
  var nodeName=tab=='families'?'familyTreeNode':'systemTreeNode';
  var keyName=tab=='families'?'data-tree-discipline':'data-system-tree-discipline';
  var nodes=document.getElementsByName(nodeName);
  if(!nodes||nodes.length==0)return 'FAIL missing left trade tree';
  var root=null,rootActive=0,tradeNodes=[];
  for(var n=0;n<nodes.length;n++){
    var key=nodes[n].getAttribute(keyName)||'';
    var cls=' '+(nodes[n].className||'')+' ';
    if(key=='All'){
      if(cls.indexOf(' level1 ')>=0)root=nodes[n];
      if(cls.indexOf(' active ')>=0)rootActive++;
    }else if(cls.indexOf(' level1 ')>=0){
      tradeNodes.push(nodes[n]);
    }
  }
  if(!root)return 'FAIL missing All Trades root';
  if(rootActive!=1)return 'FAIL All Trades root is not the sole active tree target';
  var state=tab=='families'?(window.currentTreeDiscipline||''):(window.currentSystemTreeDiscipline||'');
  if(state!='All')return 'FAIL initial tree state did not reset to All';
  var rows=window.rowsFor?rowsFor(tab):[],selectedRows=0;
  for(var r=0;r<rows.length;r++){
    if((rows[r].className||'').indexOf('data')<0)continue;
    var rowKey=rows[r].getAttribute('data-tree-discipline-key')||rows[r].getAttribute('data-discipline-key')||'';
    if(rowKey==selected)selectedRows++;
  }
  if(selectedRows==0&&tradeNodes.length==0)return 'OK';
  if(tradeNodes.length!=1)return 'FAIL left tree contains '+tradeNodes.length+' trade roots for selected target';
  if((tradeNodes[0].getAttribute(keyName)||'')!=selected)return 'FAIL left tree trade key stayed on '+(tradeNodes[0].getAttribute(keyName)||'')+' instead of '+selected;
  if(directText(tradeNodes[0])!=selectedLabel)return 'FAIL left tree trade label stayed on '+directText(tradeNodes[0])+' instead of '+selectedLabel;
  var rootCounts=root.getElementsByTagName('span');
  var tradeCounts=tradeNodes[0].getElementsByTagName('span');
  var rootCount=rootCounts.length?parseInt(rootCounts[0].innerText||rootCounts[0].textContent||'0',10):0;
  var tradeCount=tradeCounts.length?parseInt(tradeCounts[0].innerText||tradeCounts[0].textContent||'0',10):0;
  if(rootCount!=tradeCount||rootCount!=selectedRows)return 'FAIL left tree counts do not match selected trade rows';
  return 'OK';
})()");
		if (!string.Equals(validation, "OK", StringComparison.OrdinalIgnoreCase) && !string.Equals(validation, "SKIP", StringComparison.OrdinalIgnoreCase))
		{
			result.Failures.Add("Browse trade tree synchronization check failed: " + validation);
		}
	}

	private static void CheckBrowseRowDisciplineSynchronization(WebBrowser browser, AuditResult result)
	{
		string validation = EvalString(browser, @"
(function(){
  var tab=window.currentTab||'';
  if(tab!='families'&&tab!='systems')return 'SKIP';
  function directText(e){var value='',nodes=e?e.childNodes:[];for(var i=0;i<nodes.length;i++)if(nodes[i].nodeType==3)value+=nodes[i].nodeValue||'';return value.replace(/^\s+|\s+$/g,'');}
  function cellText(e){return String(e?(e.innerText||e.textContent||''):'').replace(/^\s+|\s+$/g,'');}
  var chips=document.getElementsByName('disciplineFilter');
  var selected='',selectedLabel='';
  for(var i=0;i<chips.length;i++)if((' '+(chips[i].className||'')+' ').indexOf(' active ')>=0){selected=chips[i].getAttribute('data-discipline')||'';selectedLabel=directText(chips[i]);}
  if(!selected||!selectedLabel)return 'SKIP';
  var rows=window.rowsFor?window.rowsFor(tab):[],matching=0;
  for(var r=0;r<rows.length;r++){
    if((rows[r].className||'').indexOf('data')<0)continue;
    var rowKey=rows[r].getAttribute('data-tree-discipline-key')||rows[r].getAttribute('data-discipline-key')||'';
    if(rowKey!=selected)continue;
    matching++;
    var rowLabel=rows[r].getAttribute('data-discipline')||'';
    if(rowLabel!=selectedLabel)return 'FAIL row payload trade label stayed on '+rowLabel+' instead of '+selectedLabel;
  }
  if(matching==0)return 'SKIP';
  var store=window.KKYFB&&window.KKYFB._stores?window.KKYFB._stores[tab]:null;
  var hasTradeCell=tab=='systems'||!store||store.mode=='family-admin';
  if(!hasTradeCell)return 'OK';
  var table=document.getElementById(tab+'Table');
  var rendered=table?table.getElementsByTagName('tr'):[];
  for(var d=0;d<rendered.length;d++){
    if((' '+(rendered[d].className||'')+' ').indexOf(' data ')<0)continue;
    var renderedKey=rendered[d].getAttribute('data-tree-discipline-key')||rendered[d].getAttribute('data-discipline-key')||'';
    if(renderedKey!=selected)continue;
    if(!rendered[d].cells||rendered[d].cells.length<3)return 'FAIL rendered trade cell is missing';
    var renderedLabel=cellText(rendered[d].cells[2]);
    if(renderedLabel!=selectedLabel)return 'FAIL rendered trade column stayed on '+renderedLabel+' instead of '+selectedLabel;
  }
  return 'OK';
})()");
		if (!string.Equals(validation, "OK", StringComparison.OrdinalIgnoreCase) && !string.Equals(validation, "SKIP", StringComparison.OrdinalIgnoreCase))
		{
			result.Failures.Add("Browse row trade-column synchronization check failed: " + validation);
		}
	}

	private static void CheckBrowseDisciplineSwitchFeedback(WebBrowser browser, AuditResult result)
	{
		string validation = EvalString(browser, @"
(function(){
  var tab=window.currentTab||'';
  if(tab!='families'&&tab!='systems')return 'SKIP';
  if(typeof window.beginBrowseDisciplineSwitch!='function')return 'FAIL missing beginBrowseDisciplineSwitch';
  function directText(e){var value='',nodes=e?e.childNodes:[];for(var i=0;i<nodes.length;i++)if(nodes[i].nodeType==3)value+=nodes[i].nodeValue||'';return value.replace(/^\s+|\s+$/g,'');}
  var chips=document.getElementsByName('disciplineFilter');
  if(!chips||chips.length<2)return 'SKIP';
  var original='',target='';
  for(var i=0;i<chips.length;i++){
    var key=chips[i].getAttribute('data-discipline')||'';
    if((' '+(chips[i].className||'')+' ').indexOf(' active ')>=0)original=key;
  }
  for(var j=0;j<chips.length;j++){
    var candidate=chips[j].getAttribute('data-discipline')||'';
    if(candidate&&candidate!=original){target=candidate;break;}
  }
  if(!original||!target)return 'SKIP';
  var targetChip=null;
  for(var k=0;k<chips.length;k++)if((chips[k].getAttribute('data-discipline')||'')==target){targetChip=chips[k];break;}
  var targetLabel=directText(targetChip);
  var onclick=targetChip?(targetChip.getAttribute('onclick')||targetChip.onclick||''):'';
  if(String(onclick).indexOf('beginBrowseDisciplineSwitch')<0)return 'FAIL trade chip has no immediate switch feedback';
	var originalTree=window.currentTreeDiscipline||'All';
	var originalTreeGroup=window.currentTreeGroup||'';
	var originalTreeCategory=window.currentTreeCategory||'';
	var originalSystemTree=window.currentSystemTreeDiscipline||'All';
	var originalSystemCategory=window.currentSystemTreeCategory||'';
  try{
	window.currentTreeDiscipline='STALE_TRADE';window.currentTreeGroup='STALE_GROUP';window.currentTreeCategory='STALE_CATEGORY';
	window.currentSystemTreeDiscipline='STALE_TRADE';window.currentSystemTreeCategory='STALE_CATEGORY';
    window.beginBrowseDisciplineSwitch(target);
    if(window.currentDiscipline!=target)return 'FAIL currentDiscipline did not change';
    if((document.body.getAttribute('data-kkyfb-browse-switch')||'')!=target)return 'FAIL switch marker did not change';
    var active=0,activeKey='';
    for(var n=0;n<chips.length;n++)if((' '+(chips[n].className||'')+' ').indexOf(' active ')>=0){active++;activeKey=chips[n].getAttribute('data-discipline')||'';}
    if(active!=1||activeKey!=target)return 'FAIL active trade chip did not follow target';
	if(window.currentTreeDiscipline!='All'||window.currentTreeGroup!=''||window.currentTreeCategory!='')return 'FAIL family tree filter was not reset';
	if(window.currentSystemTreeDiscipline!='All'||window.currentSystemTreeCategory!='')return 'FAIL system tree filter was not reset';
	var treeName=tab=='families'?'familyTreeNode':'systemTreeNode';
	var treeKey=tab=='families'?'data-tree-discipline':'data-system-tree-discipline';
	var treeNodes=document.getElementsByName(treeName),treeActive=0,treeActiveKey='',tradeLabelNodes=0;
	for(var q=0;q<treeNodes.length;q++){
		var treeNodeKey=treeNodes[q].getAttribute(treeKey)||'',treeNodeClass=' '+(treeNodes[q].className||'')+' ';
		if(treeNodeClass.indexOf(' active ')>=0){treeActive++;treeActiveKey=treeNodeKey;}
		if(treeNodeKey!='All'&&treeNodeClass.indexOf(' level1 ')>=0){tradeLabelNodes++;if(directText(treeNodes[q])!=targetLabel)return 'FAIL left tree trade label did not follow '+targetLabel;}
	}
	if(treeActive!=1||treeActiveKey!='All')return 'FAIL left tree active state did not reset to All Trades';
    var rows=window.rowsFor?rowsFor(tab):[];
    for(var r=0;r<rows.length;r++){
      if((rows[r].className||'').indexOf('data')<0||rows[r].style.display=='none')continue;
      var rowKey=rows[r].getAttribute('data-tree-discipline-key')||rows[r].getAttribute('data-discipline-key')||'';
      if(rowKey&&rowKey!=target)return 'FAIL previous-trade row remained visible during switch';
    }
    return 'OK';
  }finally{
    window.setDisciplineFilter(original);
	if(typeof window.syncBrowseTreeTradeLabel=='function')window.syncBrowseTreeTradeLabel(original);
	window.currentTreeDiscipline=originalTree;window.currentTreeGroup=originalTreeGroup;window.currentTreeCategory=originalTreeCategory;
	window.currentSystemTreeDiscipline=originalSystemTree;window.currentSystemTreeCategory=originalSystemCategory;
	if(typeof window.updateFamilyTreeActive=='function')window.updateFamilyTreeActive();
	if(typeof window.updateSystemTreeActive=='function')window.updateSystemTreeActive();
    try{document.body.removeAttribute('data-kkyfb-browse-switch');document.body.removeAttribute('data-kkyfb-browse-switch-label');}catch(e){}
  }
})()");
		if (!string.Equals(validation, "OK", StringComparison.OrdinalIgnoreCase) && !string.Equals(validation, "SKIP", StringComparison.OrdinalIgnoreCase))
		{
			result.Failures.Add("Browse trade switch feedback check failed: " + validation);
		}
	}

	private static void CaptureDashboardVisualHtml(WebBrowser browser, Dictionary<string, string> options, AuditResult result)
	{
		string outputPath = GetOption(options, "visualHtmlOut", string.Empty);
		if (string.IsNullOrWhiteSpace(outputPath))
		{
			return;
		}
		try
		{
			string standaloneHtml = EvalString(browser, @"
(function(){
  try{
    var clone=document.documentElement.cloneNode(true);var scripts=clone.getElementsByTagName('script');
    for(var i=scripts.length-1;i>=0;i--){if(scripts[i].parentNode)scripts[i].parentNode.removeChild(scripts[i]);}
    var body=clone.getElementsByTagName('body')[0];if(body){body.style.overflow='hidden';body.removeAttribute('data-kkyfb-audit-overflow');}
    return '<!doctype html>'+clone.outerHTML;
  }catch(e){return '';}
})()");
			if (string.IsNullOrWhiteSpace(standaloneHtml) || standaloneHtml.IndexOf("dashboardRoot", StringComparison.OrdinalIgnoreCase) < 0)
			{
				result.Failures.Add("Standalone dashboard visual HTML could not be exported.");
				return;
			}
			outputPath = Path.GetFullPath(outputPath);
			Directory.CreateDirectory(Path.GetDirectoryName(outputPath));
			File.WriteAllText(outputPath, standaloneHtml, new UTF8Encoding(false));
		}
		catch (Exception ex)
		{
			result.Failures.Add("Standalone dashboard visual HTML export failed: " + ex.Message);
		}
	}

	private static void CaptureBrowserScreenshot(WebBrowser browser, Dictionary<string, string> options, AuditResult result)
	{
		string outputPath = GetOption(options, "screenshotOut", string.Empty);
		if (string.IsNullOrWhiteSpace(outputPath))
		{
			return;
		}
		CaptureBrowserScreenshotToPath(browser, outputPath, result, "WebBrowser");
	}

	private static void CaptureDetailScreenshot(WebBrowser browser, Dictionary<string, string> options, AuditResult result)
	{
		string outputPath = GetOption(options, "detailScreenshotOut", string.Empty);
		string htmlOutputPath = GetOption(options, "detailHtmlOut", string.Empty);
		if (string.IsNullOrWhiteSpace(outputPath) && string.IsNullOrWhiteSpace(htmlOutputPath))
		{
			return;
		}
		string detailWindowTitle = string.Equals(GetOption(options, "languageCode", "ko"), "en", StringComparison.OrdinalIgnoreCase)
			? "Item Detail"
			: "상세 항목";
		string prepared = EvalString(browser, @"
(function(){
  try{
    var body=document.body,panel=document.getElementById('selectionDetailPanel');
    if(!body||!panel)return 'missing-detail-panel';
    var old=document.getElementById('kkyfbAuditDetailCapture');if(old&&old.parentNode)old.parentNode.removeChild(old);
    var overlay=document.createElement('div');overlay.id='kkyfbAuditDetailCapture';overlay.className='detached-shell';
    overlay.style.position='absolute';overlay.style.left='0';overlay.style.right='0';overlay.style.top='0';overlay.style.minHeight='100%';overlay.style.zIndex='100000';overlay.style.boxSizing='border-box';
    var head=document.createElement('div');head.className='detached-top';
    var kicker=document.createElement('div');kicker.className='detached-kicker';kicker.innerText='KKY Family Browser';head.appendChild(kicker);
    var title=document.createElement('h1');title.innerText=" + JsStringLiteral(detailWindowTitle) + @";head.appendChild(title);
    var kind=document.createElement('div');kind.className='detached-kind';var kindSource=document.getElementById('detailKind');kind.innerText=kindSource?(kindSource.innerText||kindSource.textContent||''):'Item Detail';head.appendChild(kind);
    var content=document.createElement('div');content.className='detached-content';content.innerHTML=panel.innerHTML;
    overlay.appendChild(head);overlay.appendChild(content);
    var children=[];for(var i=0;i<body.children.length;i++)children.push(body.children[i]);
    for(var j=0;j<children.length;j++){var child=children[j];child.setAttribute('data-kkyfb-audit-display',child.style.display||'');child.style.display='none';}
    body.setAttribute('data-kkyfb-audit-overflow',body.style.overflow||'');body.style.overflow='auto';body.appendChild(overlay);
    body.className=(body.className||'')+' fb-detail-capture';
    if(window.fitVisiblePreviewImages)setTimeout(fitVisiblePreviewImages,20);
    return 'OK';
  }catch(e){return 'ERR '+(e.message||e.description||e);}
})()");
		if (!string.Equals(prepared, "OK", StringComparison.OrdinalIgnoreCase))
		{
			result.Failures.Add("Detail screenshot layout could not be prepared: " + prepared);
			return;
		}
		string detailVisual = EvalString(browser, @"
(function(){
  function css(el,name){var s=window.getComputedStyle?window.getComputedStyle(el,null):el.currentStyle;return s?(s[name]||''):'';}
  var root=document.getElementById('kkyfbAuditDetailCapture');if(!root)return 'missing-root';
  var h=root.querySelector?root.querySelector('.detached-top h1'):null,n=root.querySelector?root.querySelector('.detail-name'):null;
  var ht=h?(h.innerText||h.textContent||'').replace(/\s+/g,''):'';var nt=n?(n.innerText||n.textContent||'').replace(/\s+/g,''):'';
  if(ht&&nt&&ht==nt)return 'duplicate-title';
  if(window.currentTab=='systems'){
    var legacyPreview=root.querySelector?root.querySelector('#previewBlock'):null;
    var nestedBlock=root.querySelector?root.querySelector('#detailNestedBlock'):null;
    if(legacyPreview&&css(legacyPreview,'display')!='none')return 'system-legacy-preview-visible';
    if(nestedBlock&&css(nestedBlock,'display')!='none')return 'system-family-composition-visible';
  }
  var old={'rgb(15, 112, 78)':1,'rgb(24, 115, 79)':1,'rgb(31, 122, 88)':1};var actionSelectors=['.fingerprint-diff-toggle','.preview-open-chip','.parameter-formula-open'];
  for(var i=0;i<actionSelectors.length;i++){var actions=root.querySelectorAll?root.querySelectorAll(actionSelectors[i]):[];for(var j=0;j<actions.length;j++){var bg=css(actions[j],'backgroundColor');if(old[bg])return 'legacy-action:'+actionSelectors[i]+'='+bg;}}
  if(document.body.getAttribute('data-theme')=='dark'){
    var surfaceSelectors=['.nested-child-list','.family-type-panel','.parameter-type-panel','.parameter-panel','.fingerprint-diff-wrap','.fingerprint-diff-head','.fingerprint-diff-summary','.family-type-table-wrap'];
    for(var s=0;s<surfaceSelectors.length;s++){var nodes=root.querySelectorAll?root.querySelectorAll(surfaceSelectors[s]):[];for(var k=0;k<nodes.length;k++){var bg=css(nodes[k],'backgroundColor');var m=(bg||'').match(/rgba?\((\d+),\s*(\d+),\s*(\d+)/);if(m&&parseInt(m[1],10)>225&&parseInt(m[2],10)>225&&parseInt(m[3],10)>225)return 'light-surface:'+surfaceSelectors[s]+'='+bg;}}
  }
  return 'OK';
})()");
		if (!string.Equals(detailVisual, "OK", StringComparison.OrdinalIgnoreCase))
		{
			result.Failures.Add("Detached detail theme check failed: " + detailVisual);
		}
		if (!string.IsNullOrWhiteSpace(htmlOutputPath))
		{
			try
			{
				string standaloneHtml = EvalString(browser, @"
(function(){
  try{
    var clone=document.documentElement.cloneNode(true);var scripts=clone.getElementsByTagName('script');
    for(var i=scripts.length-1;i>=0;i--){if(scripts[i].parentNode)scripts[i].parentNode.removeChild(scripts[i]);}
    var clonedBody=clone.getElementsByTagName('body')[0];if(clonedBody){var bodyChildren=[];for(var b=0;b<clonedBody.children.length;b++)bodyChildren.push(clonedBody.children[b]);for(var c=0;c<bodyChildren.length;c++){if(bodyChildren[c].id!='kkyfbAuditDetailCapture'&&bodyChildren[c].parentNode)bodyChildren[c].parentNode.removeChild(bodyChildren[c]);}clonedBody.style.width='';clonedBody.style.height='';clonedBody.style.overflow='auto';clonedBody.removeAttribute('data-kkyfb-audit-overflow');}
    return '<!doctype html>'+clone.outerHTML;
  }catch(e){return '';}
})()");
				if (string.IsNullOrWhiteSpace(standaloneHtml) || standaloneHtml.IndexOf("kkyfbAuditDetailCapture", StringComparison.OrdinalIgnoreCase) < 0)
				{
					result.Failures.Add("Standalone detached detail HTML could not be exported.");
				}
				else
				{
					htmlOutputPath = Path.GetFullPath(htmlOutputPath);
					Directory.CreateDirectory(Path.GetDirectoryName(htmlOutputPath));
					File.WriteAllText(htmlOutputPath, standaloneHtml, new UTF8Encoding(false));
				}
			}
			catch (Exception ex)
			{
				result.Failures.Add("Standalone detached detail HTML export failed: " + ex.Message);
			}
		}
		try
		{
			Form host = browser.FindForm();
			Size previousSize = host == null ? Size.Empty : host.ClientSize;
			if (host != null)
			{
				host.ClientSize = new Size(1180, 900);
			}
			DoEventsFor(220);
			if (!string.IsNullOrWhiteSpace(outputPath))
			{
				CaptureBrowserScreenshotToPath(browser, outputPath, result, "Detail WebBrowser");
			}
			if (host != null)
			{
				host.ClientSize = previousSize;
			}
		}
		finally
		{
			EvalString(browser, @"
(function(){
  try{
    var body=document.body,overlay=document.getElementById('kkyfbAuditDetailCapture');if(overlay&&overlay.parentNode)overlay.parentNode.removeChild(overlay);
    var children=body?body.children:[];for(var i=0;i<children.length;i++){var child=children[i];var saved=child.getAttribute('data-kkyfb-audit-display');if(saved!==null){child.style.display=saved;child.removeAttribute('data-kkyfb-audit-display');}}
    if(body){body.style.overflow=body.getAttribute('data-kkyfb-audit-overflow')||'';body.removeAttribute('data-kkyfb-audit-overflow');body.className=(body.className||'').replace(/\s*fb-detail-capture/g,'');}
    return 'OK';
  }catch(e){return 'ERR';}
})()");
			DoEventsFor(80);
		}
	}

	private static void CaptureBrowserScreenshotToPath(WebBrowser browser, string outputPath, AuditResult result, string label)
	{
		try
		{
			outputPath = Path.GetFullPath(outputPath);
			Directory.CreateDirectory(Path.GetDirectoryName(outputPath));
			int width = Math.Max(1, browser.ClientSize.Width);
			int height = Math.Max(1, browser.ClientSize.Height);
			browser.Invalidate();
			browser.Update();
			DoEventsFor(120);
			using (Bitmap bitmap = new Bitmap(width, height))
			{
				int score = 0;
				try
				{
					browser.DrawToBitmap(bitmap, new Rectangle(0, 0, width, height));
					score = VisualPixelScore(bitmap);
				}
				catch
				{
					score = 0;
				}
				if (score < 4500 || !HasVisualPixelVariance(bitmap))
				{
					result.Failures.Add(label + " screenshot was blank or could not be captured: " + outputPath + " (score=" + score + ").");
					return;
				}
				bitmap.Save(outputPath, System.Drawing.Imaging.ImageFormat.Png);
			}
		}
		catch (Exception ex)
		{
			result.Failures.Add(label + " screenshot failed: " + ex.Message);
		}
	}

	private static int VisualPixelScore(Bitmap bitmap)
	{
		if (bitmap == null || bitmap.Width < 2 || bitmap.Height < 2)
		{
			return 0;
		}
		HashSet<int> colors = new HashSet<int>();
		int informative = 0;
		int stepX = Math.Max(1, bitmap.Width / 48);
		int stepY = Math.Max(1, bitmap.Height / 32);
		for (int y = 0; y < bitmap.Height; y += stepY)
		{
			for (int x = 0; x < bitmap.Width; x += stepX)
			{
				Color color = bitmap.GetPixel(x, y);
				int quantized = (color.R / 16 << 8) | (color.G / 16 << 4) | color.B / 16;
				colors.Add(quantized);
				int brightness = (color.R * 299 + color.G * 587 + color.B * 114) / 1000;
				if (brightness > 18 && brightness < 242)
				{
					informative++;
				}
			}
		}
		return colors.Count * 1000 + Math.Min(999, informative);
	}

	private static bool HasVisualPixelVariance(Bitmap bitmap)
	{
		if (bitmap == null || bitmap.Width < 2 || bitmap.Height < 2)
		{
			return false;
		}
		HashSet<int> colors = new HashSet<int>();
		int stepX = Math.Max(1, bitmap.Width / 24);
		int stepY = Math.Max(1, bitmap.Height / 16);
		for (int y = 0; y < bitmap.Height; y += stepY)
		{
			for (int x = 0; x < bitmap.Width; x += stepX)
			{
				colors.Add(bitmap.GetPixel(x, y).ToArgb());
				if (colors.Count >= 4)
				{
					return true;
				}
			}
		}
		return false;
	}

	private static void RunMessageBodyAudit(string html, Dictionary<string, string> options, AuditResult result)
	{
		using (Form host = new Form())
		using (WebBrowser browser = new WebBrowser())
		{
			host.Text = "KKY Family Browser Message Audit";
			host.ShowInTaskbar = false;
			host.StartPosition = FormStartPosition.Manual;
			host.Left = -32000;
			host.Top = -32000;
			host.Width = 900;
			host.Height = 560;
			browser.Dock = DockStyle.Fill;
			browser.ScriptErrorsSuppressed = true;
			browser.AllowWebBrowserDrop = false;
			browser.IsWebBrowserContextMenuEnabled = false;
			browser.WebBrowserShortcutsEnabled = false;
			host.Controls.Add(browser);
			host.Show();
			Stopwatch documentLoadSw = Stopwatch.StartNew();
			LoadHtmlIntoBrowser(browser, html, string.Empty, false);
			documentLoadSw.Stop();
			result.DocumentLoadMilliseconds = documentLoadSw.ElapsedMilliseconds;
			if (browser.Document == null || browser.Document.Body == null)
			{
				result.Failures.Add("WebBrowser did not finish loading the message HTML.");
				return;
			}

			string languageCode = GetOption(options, "languageCode", "ko");
			bool isKorean = !string.Equals(languageCode, "en", StringComparison.OrdinalIgnoreCase);
			bool automaticResult = string.Equals(GetOption(options, "messageFixture", string.Empty), "auto-result", StringComparison.OrdinalIgnoreCase);
			string expectedKind = automaticResult ? "information" : "error";
			if (!string.Equals(BodyAttribute(browser, "data-message-kind"), expectedKind, StringComparison.OrdinalIgnoreCase))
			{
				result.Failures.Add("Structured message did not expose data-message-kind=" + expectedKind + ".");
			}
			if (!string.Equals(BodyAttribute(browser, "data-message-structured"), "true", StringComparison.OrdinalIgnoreCase))
			{
				result.Failures.Add("Structured message did not expose data-message-structured=true.");
			}
			if (!string.Equals(BodyAttribute(browser, "data-dialog-shell"), "full-html", StringComparison.OrdinalIgnoreCase))
			{
				result.Failures.Add("Structured message is not hosted by the full HTML dialog shell.");
			}
			string expectedTheme = NormalizeThemeCode(GetOption(options, "themeCode", "light"));
			if (!string.Equals(BodyAttribute(browser, "data-theme"), expectedTheme, StringComparison.OrdinalIgnoreCase))
			{
				result.Failures.Add("Structured message theme mismatch. Expected " + expectedTheme + ".");
			}
			string[] requiredIds = automaticResult ? new[]
			{
				"dialogShell",
				"dialogHeader",
				"dialogTitle",
				"dialogBody",
				"dialogFooter",
				"dialogCopyDetails",
				"dialogOpenFolder",
				"dialogCopyPath",
				"dialogAccept",
				"messageHeadline",
				"messageResultGroups",
				"messageMetricGrid",
				"messageContextTable",
				"messageOutputList"
			} : new[]
			{
				"dialogShell",
				"dialogHeader",
				"dialogTitle",
				"dialogSubtitle",
				"dialogClose",
				"dialogBody",
				"dialogFooter",
				"dialogStatus",
				"dialogCopyDetails",
				"dialogAccept",
				"messageHeadline",
				"messageSectionCause",
				"messageSectionAction",
				"messageSectionAdmin",
				"messageSectionTechnical",
				"messageLogPath",
				"messageSupportCode",
				"messageTechnicalDetail"
			};
			foreach (string id in requiredIds)
			{
				if (browser.Document.GetElementById(id) == null)
				{
					result.Failures.Add("Structured message is missing semantic element #" + id + ".");
				}
			}
			string bodyText = browser.Document.Body.InnerText ?? string.Empty;
			if (automaticResult)
			{
				if (!string.Equals(BodyAttribute(browser, "data-message-auto-result"), "true", StringComparison.OrdinalIgnoreCase))
				{
					result.Failures.Add("Automatic result message did not expose data-message-auto-result=true.");
				}
				if (bodyText.IndexOf("AUDIT_PROJECT", StringComparison.Ordinal) < 0 || bodyText.IndexOf("AUDIT_STANDARD", StringComparison.Ordinal) < 0 || bodyText.IndexOf("24", StringComparison.Ordinal) < 0 || bodyText.IndexOf("audit-result.xlsx", StringComparison.OrdinalIgnoreCase) < 0)
				{
					result.Failures.Add("Automatic result message lost context, metric, or output-path content.");
				}
				if (isKorean && (!HangulRegex.IsMatch(bodyText) || bodyText.IndexOf("비교 결과", StringComparison.Ordinal) < 0))
				{
					result.Failures.Add("Korean automatic result labels are missing.");
				}
				if (!isKorean && HangulRegex.IsMatch(bodyText))
				{
					result.Failures.Add("English automatic result contains Hangul text.");
				}
			}
			else if (isKorean)
			{
				if (!HangulRegex.IsMatch(bodyText) || bodyText.IndexOf("실패 이유", StringComparison.Ordinal) < 0 || bodyText.IndexOf("지금 할 일", StringComparison.Ordinal) < 0)
				{
					result.Failures.Add("Korean structured message labels are missing.");
				}
			}
			else
			{
				if (HangulRegex.IsMatch(bodyText))
				{
					result.Failures.Add("English structured message contains Hangul text.");
				}
				if (bodyText.IndexOf("Why it failed", StringComparison.OrdinalIgnoreCase) < 0 || bodyText.IndexOf("What to do now", StringComparison.OrdinalIgnoreCase) < 0)
				{
					result.Failures.Add("English structured message labels are missing.");
				}
			}
			int allowedWidth = browser.ClientSize.Width + 24;
			if (browser.Document.Body.ScrollRectangle.Width > allowedWidth)
			{
				result.Failures.Add("Structured message body overflows horizontally: body=" + browser.Document.Body.ScrollRectangle.Width + ", client=" + browser.ClientSize.Width + ".");
			}
			List<Clickable> clickables = CollectClickables(browser);
			result.ClickableCount = clickables.Count;
			if (clickables.Count < 3)
			{
				result.Failures.Add("Full HTML dialog shell is missing visible close/copy/accept actions.");
			}
			browser.Stop();
			browser.DocumentText = "<html><body></body></html>";
			host.Close();
		}
	}

	private static void RunSyntheticCacheAudit(Dictionary<string, string> options, AuditResult result)
	{
		Assembly assembly = Assembly.LoadFrom(result.HostAssembly);
		Type loaderType = assembly.GetType("FamilyBrowserDataLoader", true);
		MethodInfo audit = loaderType.GetMethod("RunSyntheticPerformanceAudit", BindingFlags.Public | BindingFlags.Static);
		if (audit == null)
		{
			throw new MissingMethodException("FamilyBrowserDataLoader", "RunSyntheticPerformanceAudit");
		}
		int familyCount = ParseInt(GetOption(options, "syntheticFamilyCount", "1000"), 1000);
		int systemCount = ParseInt(GetOption(options, "syntheticSystemCount", "1000"), 1000);
		object auditResult = audit.Invoke(null, new object[] { familyCount, systemCount });
		if (auditResult == null)
		{
			result.Failures.Add("Synthetic row-cache audit returned no result.");
			return;
		}
		Type resultType = auditResult.GetType();
		bool success = Convert.ToBoolean(resultType.GetProperty("Success").GetValue(auditResult, null));
		result.CacheBytes = Convert.ToInt64(resultType.GetProperty("CacheBytes").GetValue(auditResult, null));
		result.CacheSaveMilliseconds = Convert.ToInt64(resultType.GetProperty("SaveMilliseconds").GetValue(auditResult, null));
		result.CacheColdLoadMilliseconds = Convert.ToInt64(resultType.GetProperty("ColdLoadMilliseconds").GetValue(auditResult, null));
		result.CacheWarmLoadMilliseconds = Convert.ToInt64(resultType.GetProperty("WarmLoadMilliseconds").GetValue(auditResult, null));
		result.CacheOfflineLoadMilliseconds = Convert.ToInt64(resultType.GetProperty("OfflineLoadMilliseconds").GetValue(auditResult, null));
		string error = Convert.ToString(resultType.GetProperty("ErrorMessage").GetValue(auditResult, null)) ?? string.Empty;
		if (!success)
		{
			result.Failures.Add("Synthetic row-cache audit failed: " + error);
		}
		int cacheTargetMs = ParseInt(GetOption(options, "cacheTargetMs", "1500"), 1500);
		if (result.CacheColdLoadMilliseconds > cacheTargetMs || result.CacheWarmLoadMilliseconds > cacheTargetMs || result.CacheOfflineLoadMilliseconds > cacheTargetMs)
		{
			result.Failures.Add("Synthetic row-cache load exceeded " + cacheTargetMs + "ms: cold=" + result.CacheColdLoadMilliseconds + " warm=" + result.CacheWarmLoadMilliseconds + " offline=" + result.CacheOfflineLoadMilliseconds + ".");
		}
	}

	private static void CheckPerformance(WebBrowser browser, Dictionary<string, string> options, AuditResult result)
	{
		try
		{
			string rowState = EvalString(browser, "(function(){var s=(window.KKYFB&&KKYFB.stats)?KKYFB.stats(currentTab):null;if(s)return s.total+'|'+s.rendered+'|'+s.visible+'|'+s.filtered+'|'+s.page+'|'+s.checked;var rows=window.rowsFor?rowsFor(currentTab):[];var total=0,visible=0;for(var i=0;i<rows.length;i++){if((rows[i].className||'').indexOf('data')<0)continue;total++;if(rows[i].style.display!='none')visible++;}return total+'|'+total+'|'+visible+'|'+visible+'|0|0';})()");
			string[] parts = rowState.Split('|');
			if (parts.Length >= 3)
			{
				result.DataRowCount = ParseInt(parts[0], 0);
				result.DomRowCount = ParseInt(parts[1], 0);
				result.VisibleRowCount = ParseInt(parts[2], 0);
			}
			else
			{
				result.Failures.Add("Virtual row statistics unavailable: " + rowState + "; " + EvalString(browser, "'currentTab='+(typeof currentTab)+';KKYFB='+(typeof window.KKYFB)+';stats='+(window.KKYFB?typeof KKYFB.stats:'-')+';error='+(document.body?document.body.getAttribute('data-kkyfb-js-error'):'')"));
			}
			string activeTab = EvalString(browser, "window.currentTab||''");
			int expectedRows = string.Equals(activeTab, "systems", StringComparison.OrdinalIgnoreCase)
				? ParseInt(GetOption(options, "syntheticSystemCount", "0"), 0)
				: ParseInt(GetOption(options, "syntheticFamilyCount", "0"), 0);
			if (expectedRows > 0 && result.DataRowCount != expectedRows)
			{
				result.Failures.Add("Virtual row store count mismatch: expected " + expectedRows + ", actual " + result.DataRowCount + ".");
			}
			if (result.DomRowCount > 150 || result.VisibleRowCount > 150)
			{
				result.Failures.Add("Virtual row window exceeds 150 DOM rows: DOM=" + result.DomRowCount + ", visible=" + result.VisibleRowCount + ", total=" + result.DataRowCount + ".");
			}
			if (result.DataRowCount > 150 && result.DomRowCount >= result.DataRowCount)
			{
				result.Failures.Add("All rows still exist in the DOM; true row virtualization is not active: DOM=" + result.DomRowCount + ", total=" + result.DataRowCount + ".");
			}
			if (result.DataRowCount > 150)
			{
				CheckVirtualRowPagingAndSelection(browser, activeTab, result);
			}

			HtmlElement search = browser.Document == null ? null : browser.Document.GetElementById("searchBox");
			long worst = 0L;
			string[] querySequence = new string[5] { "A", "AU", "AUD", "AUDI", "AUDIT" };
			for (int i = 0; i < querySequence.Length; i++)
			{
				if (search != null)
				{
					search.SetAttribute("value", querySequence[i]);
				}
				Stopwatch sw = Stopwatch.StartNew();
				browser.Document.InvokeScript("filterRows", new object[] { "search" });
				sw.Stop();
				worst = Math.Max(worst, sw.ElapsedMilliseconds);
			}
			result.FilterMilliseconds = worst;
			int filterTargetMs = ParseInt(GetOption(options, "filterTargetMs", "150"), 150);
			if (result.FilterMilliseconds > filterTargetMs)
			{
				result.Failures.Add("Search/filter response exceeded " + filterTargetMs + "ms: " + result.FilterMilliseconds + "ms for " + result.DataRowCount + " rows.");
			}
			int usableTargetMs = ParseInt(GetOption(options, "usableTargetMs", "1500"), 1500);
			long usableMs = result.HtmlRenderMilliseconds + result.DocumentLoadMilliseconds + result.DashboardReadyMilliseconds;
			if (usableMs > usableTargetMs)
			{
				result.Failures.Add("Synthetic local list usable time exceeded " + usableTargetMs + "ms: render=" + result.HtmlRenderMilliseconds + " document=" + result.DocumentLoadMilliseconds + " ready=" + result.DashboardReadyMilliseconds + ".");
			}
		}
		catch (Exception ex)
		{
			result.Failures.Add("Performance check threw " + DescribeException(ex));
		}
	}

	private static void CheckVirtualRowPagingAndSelection(WebBrowser browser, string activeTab, AuditResult result)
	{
		string script = @"
(function(){
  if(!window.KKYFB||!KKYFB.stats||!window.changeRowWindow||!window.goToRowWindowPage||!window.resizeColumnForAudit)return 'ERR|virtual-api-missing';
  var tab=window.currentTab||'';
  function firstEnabled(){
    var table=document.getElementById(tab+'Table');
    if(!table)return null;
    var inputs=table.getElementsByTagName('input');
    for(var i=0;i<inputs.length;i++){
      var c=' '+(inputs[i].className||'')+' ';
      if(c.indexOf(' row-check ')>=0&&!inputs[i].disabled)return inputs[i];
    }
    return null;
  }
  function toggle(box){
    if(tab=='families')return window.toggleFamilyCheck(null,box);
    return window.toggleSystemCheck(null,box);
  }
  function firstDataCells(table){
    var rows=table?table.getElementsByTagName('tr'):[];
    for(var r=0;r<rows.length;r++){
      if((' '+(rows[r].className||'')+' ').indexOf(' data ')>=0)return rows[r].getElementsByTagName('td');
    }
    return [];
  }
  function offsets(cells){
    var values=[];
    for(var i=0;i<cells.length;i++)values.push(cells[i].offsetWidth);
    return values;
  }
  var first=firstEnabled();
  if(!first)return 'ERR|first-page-checkbox-missing';
  var fixed=document.getElementById(tab+'FixedHeader');
  var headTable=fixed&&fixed.getElementsByTagName('table').length?fixed.getElementsByTagName('table')[0]:null;
  var bodyTable=document.getElementById(tab+'Table');
  var headCols=headTable?headTable.getElementsByTagName('col'):[];
  var bodyCols=bodyTable?bodyTable.getElementsByTagName('col'):[];
  var handles=fixed?fixed.getElementsByTagName('span'):[];
  var handleCount=0;
  for(var h=0;h<handles.length;h++)if((' '+(handles[h].className||'')+' ').indexOf(' column-resize-handle ')>=0)handleCount++;
  var widths=KKYFB.columnWidths?KKYFB.columnWidths(tab):[];
  var resizeIndex=widths.length>1?1:0;
  var minimum=resizeIndex===0?42:72;
  var resizeTarget=widths[resizeIndex]>minimum+24?widths[resizeIndex]-23:widths[resizeIndex]+37;
  var headCells=headTable?headTable.getElementsByTagName('th'):[];
  var bodyCells=firstDataCells(bodyTable);
  var headOffsets=offsets(headCells);
  var bodyOffsets=offsets(bodyCells);
  var resized=resizeColumnForAudit(tab,resizeIndex,resizeTarget);
  var resizedWidths=KKYFB.columnWidths?KKYFB.columnWidths(tab):[];
  var resizeSynced=!!(resized&&headCols.length>resizeIndex&&bodyCols.length>resizeIndex&&parseInt(headCols[resizeIndex].style.width,10)==resizeTarget&&parseInt(bodyCols[resizeIndex].style.width,10)==resizeTarget&&resizedWidths[resizeIndex]==resizeTarget);
  var untouchedStable=true;
  for(var c=0;c<widths.length;c++){
    if(c==resizeIndex)continue;
    if(resizedWidths[c]!=widths[c]||parseInt(headCols[c].style.width,10)!=widths[c]||parseInt(bodyCols[c].style.width,10)!=widths[c])untouchedStable=false;
    if(headCells[c]&&Math.abs(headCells[c].offsetWidth-headOffsets[c])>1)untouchedStable=false;
    if(bodyCells[c]&&Math.abs(bodyCells[c].offsetWidth-bodyOffsets[c])>1)untouchedStable=false;
  }
  var resizedTotal=0;
  for(var t=0;t<resizedWidths.length;t++)resizedTotal+=parseInt(resizedWidths[t],10)||0;
  var headLocked=headTable&&(' '+(headTable.className||'')+' ').indexOf(' kkyfb-column-width-locked ')>=0;
  var bodyLocked=bodyTable&&(' '+(bodyTable.className||'')+' ').indexOf(' kkyfb-column-width-locked ')>=0;
  var exactTableWidth=!!(headLocked&&bodyLocked&&parseInt(headTable.style.width,10)==resizedTotal&&parseInt(bodyTable.style.width,10)==resizedTotal&&Math.abs(headTable.offsetWidth-resizedTotal)<=2&&Math.abs(bodyTable.offsetWidth-resizedTotal)<=2);
  var firstKey=first.getAttribute('data-row-key')||'';
  first.checked=true;toggle(first);
  goToRowWindowPage(1);
  var pageTwo=KKYFB.stats(tab);
  var pageHost=document.getElementById('rowWindowPages');
  var pageLinks=pageHost?pageHost.getElementsByTagName('a'):[];
  var activePage=0;
  for(var p=0;p<pageLinks.length;p++)if(pageLinks[p].getAttribute('data-page-index')=='1'&&(' '+(pageLinks[p].className||'')+' ').indexOf(' active ')>=0)activePage=1;
  var pageSummary=document.getElementById('rowWindowPageSummary');
  var pageSummaryOnTwo=!!(pageSummary&&pageSummary.innerText.indexOf('2')>=0);
  var second=firstEnabled();
  if(!second)return 'ERR|second-page-checkbox-missing';
  var secondKey=second.getAttribute('data-row-key')||'';
  second.checked=true;toggle(second);
  var checked=(tab=='families'?checkedFamilyRows():checkedSystemRows()).length;
  changeRowWindow(-1);
  var restored=firstEnabled();
  var persisted=!!(restored&&restored.checked);
  var returned=KKYFB.stats(tab);
  if(tab=='families')clearFamilySelection();else clearSystemSelection();
	var cleared=KKYFB.stats(tab);
	var selectedAfterClear=0;
	var table=document.getElementById(tab+'Table');
	var rows=table?table.getElementsByTagName('tr'):[];
	for(var i=0;i<rows.length;i++)if((' '+(rows[i].className||'')+' ').indexOf(' selected ')>=0)selectedAfterClear++;
  return 'OK|'+pageTwo.page+'|'+checked+'|'+(persisted?'1':'0')+'|'+pageTwo.rendered+'|'+(firstKey!=secondKey?'1':'0')+'|'+returned.page+'|'+cleared.checked+'|'+selectedAfterClear+'|'+(pageLinks.length>=2?'1':'0')+'|'+activePage+'|'+(pageSummaryOnTwo?'1':'0')+'|'+(handleCount==headCols.length?'1':'0')+'|'+(resizeSynced?'1':'0')+'|'+(untouchedStable?'1':'0')+'|'+(exactTableWidth?'1':'0');
})()
";
		string state = EvalString(browser, script);
		string[] parts = state.Split('|');
		if (parts.Length < 16 || !string.Equals(parts[0], "OK", StringComparison.Ordinal))
		{
			result.Failures.Add("Virtual row paging smoke failed for " + activeTab + ": " + state + ".");
			return;
		}
		if (ParseInt(parts[1], -1) != 1 || ParseInt(parts[2], 0) != 2 || parts[3] != "1" || ParseInt(parts[4], 0) > 150 || parts[5] != "1" || ParseInt(parts[6], -1) != 0 || ParseInt(parts[7], -1) != 0 || ParseInt(parts[8], -1) != 0 || parts[9] != "1" || parts[10] != "1" || parts[11] != "1" || parts[12] != "1" || parts[13] != "1")
		{
			result.Failures.Add("Virtual row paging, numbered navigation, column resize, or cross-page selection state is invalid for " + activeTab + ": " + state + ".");
		}
		if (parts[14] != "1" || parts[15] != "1")
		{
			result.Failures.Add("Column resize changed an untouched column or failed to lock the exact table width for " + activeTab + ": " + state + ".");
		}
	}

	private static void EvaluateInlineScripts(WebBrowser browser, string html, AuditResult result)
	{
		if (string.IsNullOrWhiteSpace(html) || browser.Document == null)
		{
			return;
		}

		MatchCollection matches = Regex.Matches(html, @"<script\b[^>]*>([\s\S]*?)</script>", RegexOptions.IgnoreCase);
		int count = 0;
		foreach (Match match in matches)
		{
			string script = match.Groups.Count > 1 ? match.Groups[1].Value : string.Empty;
			if (string.IsNullOrWhiteSpace(script))
			{
				continue;
			}
			try
			{
				browser.Document.InvokeScript("eval", new object[] { script });
				count++;
			}
			catch (Exception ex)
			{
				result.Failures.Add("Inline dashboard script eval failed: " + DescribeException(ex));
				break;
			}
		}

		if (count > 0)
		{
			result.Warnings.Add("Evaluated " + count + " inline dashboard script block(s) because the WebBrowser control had not executed them before the audit click pass.");
		}
	}

	private static bool LoadHtmlIntoBrowser(WebBrowser browser, string html, string htmlPath, bool navigateToFile)
	{
		bool completed = false;
		WebBrowserDocumentCompletedEventHandler handler = delegate(object sender, WebBrowserDocumentCompletedEventArgs e)
		{
			if (browser.ReadyState == WebBrowserReadyState.Complete)
			{
				completed = true;
			}
		};
		browser.DocumentCompleted += handler;
		try
		{
			if (navigateToFile)
			{
				if (string.IsNullOrWhiteSpace(htmlPath))
				{
					throw new ArgumentException("HTML path is required for file navigation.", "htmlPath");
				}
				browser.Navigate(new Uri(Path.GetFullPath(htmlPath)));
			}
			else
			{
				browser.DocumentText = html ?? string.Empty;
			}
			return WaitUntil(() => completed && browser.Document != null && browser.Document.Body != null, 15000);
		}
		finally
		{
			browser.DocumentCompleted -= handler;
		}
	}

	private static bool WaitForDashboardReady(WebBrowser browser, int timeoutMs)
	{
		return WaitUntil(() => IsDashboardReady(browser), timeoutMs);
	}

	private static bool IsDashboardReady(WebBrowser browser)
	{
		if (BodyAttribute(browser, "data-kkyfb-ready") == "true")
		{
			return true;
		}
		string diagnostic = DashboardReadyDiagnostic(browser);
		return diagnostic.IndexOf("openAdvancedFilter=function", StringComparison.OrdinalIgnoreCase) >= 0
			&& diagnostic.IndexOf("setFilter=function", StringComparison.OrdinalIgnoreCase) >= 0
			&& diagnostic.IndexOf("loadFamilySelection=function", StringComparison.OrdinalIgnoreCase) >= 0
			&& diagnostic.IndexOf("prepareDetailWindowOpen=function", StringComparison.OrdinalIgnoreCase) >= 0;
	}

	private static void ForceDashboardOnLoad(WebBrowser browser, AuditResult result)
	{
		string script = @"
(function(){
  try{
    if(window.onload)window.onload();
    if(document.body&&!document.body.getAttribute('data-kkyfb-ready'))document.body.setAttribute('data-kkyfb-ready','forced');
    return 'OK';
  }catch(e){
    try{if(document.body)document.body.setAttribute('data-kkyfb-js-error','onload '+(e.message||e.description||e));}catch(e2){}
    return 'ERR '+(e.message||e.description||e);
  }
})()
";
		string value = EvalString(browser, script);
		if (!string.IsNullOrWhiteSpace(value) && value.StartsWith("ERR ", StringComparison.OrdinalIgnoreCase))
		{
			result.Warnings.Add("Forced dashboard onload reported: " + value);
		}
	}

	private static string DashboardReadyDiagnostic(WebBrowser browser)
	{
		string script = @"
(function(){
  function t(n){try{return n+'='+typeof window[n];}catch(e){return n+'=ERR';}}
  var bodyReady='';
  try{bodyReady=document.body?document.body.getAttribute('data-kkyfb-ready'):'';}catch(e2){}
  return 'ready='+bodyReady+
    ';state='+document.readyState+
    ';'+t('openAdvancedFilter')+
    ';'+t('setFilter')+
    ';'+t('loadFamilySelection')+
    ';'+t('prepareDetailWindowOpen')+
    ';'+t('toggleFamilyTree')+
    ';error='+(document.body?document.body.getAttribute('data-kkyfb-js-error'):'');
})()
";
		return EvalString(browser, script);
	}

	private static string EvalString(WebBrowser browser, string script)
	{
		try
		{
			return Convert.ToString(browser.Document.InvokeScript("eval", new object[] { script })) ?? string.Empty;
		}
		catch (Exception ex)
		{
			return "eval-error:" + ex.GetType().Name + ":" + ex.Message;
		}
	}

	private static int ParseInt(string value, int fallback)
	{
		int parsed;
		return int.TryParse(value, out parsed) ? parsed : fallback;
	}

	private static void InjectAuditHooks(WebBrowser browser, AuditResult result)
	{
		string script = @"
window.alert=function(m){try{document.body.setAttribute('data-kkyfb-last-alert', String(m||''));}catch(e){} return true;};
window.confirm=function(m){try{document.body.setAttribute('data-kkyfb-last-confirm', String(m||''));}catch(e){} return false;};
window.onerror=function(msg,url,line){try{document.body.setAttribute('data-kkyfb-js-error', String(msg||'')+' line '+line);}catch(e){} return true;};
try{document.body.setAttribute('data-kkyfb-audit-hooks','true');}catch(e){}
";
		try
		{
			browser.Document.InvokeScript("eval", new object[] { script });
		}
		catch (Exception ex)
		{
			result.Failures.Add("Could not inject audit JS hooks: " + ex.Message);
		}
	}

	private static void CheckLayout(WebBrowser browser, AuditResult result)
	{
		string script = @"
(function(){
  function first(cls){var x=document.getElementsByClassName(cls);return x&&x.length?x[0]:null;}
  function rect(el){if(!el||!el.getBoundingClientRect)return null;var r=el.getBoundingClientRect();return {l:r.left,t:r.top,r:r.right,b:r.bottom,w:r.right-r.left,h:r.bottom-r.top};}
  var toolbar=rect(first('toolbar'));
  var layout=rect(first('layout'));
  var top=rect(first('top'));
  var pills=rect(first('pills'));
  var actions=rect(first('top-actions'));
  var context=rect(first('project-context'));
  var out=[];
  if(!toolbar)out.push('missing toolbar');
  if(!layout)out.push('missing layout');
  if(!actions)out.push('missing top action buttons');
  if(!context)out.push('missing project file context');
  if(toolbar&&layout&&layout.l+1<toolbar.r)out.push('layout overlaps toolbar: layout.left='+layout.l+', toolbar.right='+toolbar.r);
  if(top&&layout&&layout.t+1<top.b)out.push('layout overlaps top: layout.top='+layout.t+', top.bottom='+top.b);
  if(top&&pills&&pills.b>top.b+4)out.push('pills overflow top: pills.bottom='+pills.b+', top.bottom='+top.b);
  if(actions&&context&&context.t+1<actions.b)out.push('project context overlaps top actions: context.top='+context.t+', actions.bottom='+actions.b);
  if(context&&pills&&pills.t<context.b+12)out.push('status pills too close to project context: pills.top='+pills.t+', context.bottom='+context.b);
  if(actions&&pills&&pills.t<actions.b+24)out.push('status pills too close to top actions: pills.top='+pills.t+', actions.bottom='+actions.b);
  function css(el,name){var s=(window.getComputedStyle?window.getComputedStyle(el,null):el.currentStyle);return s?(s[name]||s.getPropertyValue&&s.getPropertyValue(name)||''):'';}
  function checkFilters(cls){
    var bars=document.getElementsByClassName?document.getElementsByClassName(cls):[];
    for(var i=0;i<bars.length;i++){
      var links=bars[i].getElementsByTagName('a');
      for(var j=0;j<links.length;j++){
        var el=links[j], c=' '+(el.className||' ')+' ';
        if(c.indexOf(' filter ')<0)continue;
        var label=(el.innerText||el.textContent||'').replace(/\s+/g,' ').replace(/^\s+|\s+$/g,'');
        var overflow=(css(el,'overflow')||css(el,'overflowX')||'').toLowerCase();
        var textOverflow=(css(el,'textOverflow')||css(el,'text-overflow')||'').toLowerCase();
        if(textOverflow.indexOf('ellipsis')>=0)out.push(cls+' filter still uses ellipsis: '+label);
        if(overflow.indexOf('hidden')>=0 && el.scrollWidth>el.clientWidth+2)out.push(cls+' filter content is clipped: '+label);
      }
    }
  }
  function checkInlineStatusFilters(){
    var groups=document.getElementsByClassName?document.getElementsByClassName('inline-status-toggle'):[];
    for(var i=0;i<groups.length;i++){
      var group=groups[i];
      var groupOverflow=(css(group,'overflow')||css(group,'overflowX')||'').toLowerCase();
      var groupTextOverflow=(css(group,'textOverflow')||css(group,'text-overflow')||'').toLowerCase();
      if(groupTextOverflow.indexOf('ellipsis')>=0)out.push('inline status group still uses ellipsis');
      if(groupOverflow.indexOf('hidden')>=0 && group.scrollWidth>group.clientWidth+2)out.push('inline status group content is clipped');
      var links=group.getElementsByTagName('a');
      for(var j=0;j<links.length;j++){
        var el=links[j], c=' '+(el.className||' ')+' ';
        if(c.indexOf(' inline-status-filter ')<0)continue;
        var label=(el.innerText||el.textContent||'').replace(/\s+/g,' ').replace(/^\s+|\s+$/g,'');
        var overflow=(css(el,'overflow')||css(el,'overflowX')||'').toLowerCase();
        var textOverflow=(css(el,'textOverflow')||css(el,'text-overflow')||'').toLowerCase();
        if(textOverflow.indexOf('ellipsis')>=0)out.push('inline status filter still uses ellipsis: '+label);
        if(overflow.indexOf('hidden')>=0 && el.scrollWidth>el.clientWidth+2)out.push('inline status filter content is clipped: '+label);
      }
      var head=group.parentNode;
      while(head&&((' '+(head.className||'')+' ').indexOf(' pane-head ')<0))head=head.parentNode;
      var pane=head?head.parentNode:null;
      var grid=null;
      if(pane){
        var divs=pane.getElementsByTagName('div');
        for(var k=0;k<divs.length;k++){
          if((' '+(divs[k].className||'')+' ').indexOf(' family-browser-grid ')>=0){grid=divs[k];break;}
        }
      }
      var hr=rect(head), gr=rect(grid);
      if(hr&&gr&&gr.t+1<hr.b)out.push('browser grid overlaps action status filters: grid.top='+gr.t+', head.bottom='+hr.b);
    }
  }
  function checkPermissionDiagnosticRows(){
    var grids=document.getElementsByClassName?document.getElementsByClassName('permission-diagnostic-grid'):[];
    if(!grids||grids.length!==1){out.push('permission diagnostic grid count is '+(grids?grids.length:0));return;}
    var grid=grids[0], gr=rect(grid);
    var all=grid.getElementsByClassName?grid.getElementsByClassName('diagnostic-card'):[];
    if(!gr||!all||all.length!==3){out.push('permission diagnostic card count is '+(all?all.length:0));return;}
    var previousBottom=-1;
    for(var i=0;i<all.length;i++){
      var card=all[i], cr=rect(card);
      if(!cr)continue;
      if(Math.abs(cr.l-gr.l)>4||Math.abs(cr.r-gr.r)>4)out.push('permission diagnostic row '+(i+1)+' is not full width');
      if(previousBottom>=0&&cr.t<previousBottom+2)out.push('permission diagnostic rows '+i+' and '+(i+1)+' are not vertically stacked');
      previousBottom=cr.b;
      var titles=card.getElementsByClassName?card.getElementsByClassName('diagnostic-title'):[];
      var values=card.getElementsByClassName?card.getElementsByClassName('diagnostic-value'):[];
      var details=card.getElementsByClassName?card.getElementsByClassName('diagnostic-detail'):[];
      if(titles.length!==1||values.length!==1||details.length!==1){out.push('permission diagnostic row '+(i+1)+' has incomplete cells');continue;}
      var dr=rect(details[0]);
      if(gr.w>=900&&dr&&dr.w<320)out.push('permission diagnostic path cell is too narrow: '+dr.w);
      if(!(details[0].getAttribute('title')||''))out.push('permission diagnostic path tooltip is missing on row '+(i+1));
    }
  }
  var summaryFilterBars=document.getElementsByClassName?document.getElementsByClassName('filterbar'):[];
  if(summaryFilterBars&&summaryFilterBars.length)out.push('duplicate search summary status filter row is present');
  checkFilters('disciplinebar');
  checkInlineStatusFilters();
  function hasClass(el,cls){return (' '+((el&&el.className)||'')+' ').indexOf(' '+cls+' ')>=0;}
  function hasBodyClass(cls){return hasClass(document.body,cls);}
  function directCards(grid){
    var cards=[], kids=grid?(grid.children||grid.childNodes):[];
    for(var i=0;i<kids.length;i++){
      var el=kids[i];
      if(el&&el.nodeType==1&&hasClass(el,'settings-action-group'))cards.push(el);
    }
    return cards;
  }
  function checkTwoColumnActionCards(gridCls,label,requireFour){
    var grids=document.getElementsByClassName?document.getElementsByClassName(gridCls):[];
    if(!grids||grids.length<1){out.push(label+' grid is missing');return;}
    for(var i=0;i<grids.length;i++){
      var gr=rect(grids[i]);
      if(!gr||gr.w<900)continue;
      var cards=directCards(grids[i]);
      if(cards.length<2){out.push(label+' has fewer than two cards');continue;}
      var a=rect(cards[0]), b=rect(cards[1]);
      if(!a||!b)continue;
      if(Math.abs(a.t-b.t)>24||b.l<a.l+(a.w*0.75))out.push(label+' first two cards are not side-by-side');
      if(requireFour&&cards.length>=4){
        var c=rect(cards[2]), d=rect(cards[3]);
        if(c&&c.t<a.t+24)out.push(label+' third card did not wrap to second row');
        if(c&&d&&(Math.abs(c.t-d.t)>24||d.l<c.l+(c.w*0.75)))out.push(label+' second row cards are not side-by-side');
      }
    }
  }
  function checkActionCardButtonWrap(gridCls,label){
    var grids=document.getElementsByClassName?document.getElementsByClassName(gridCls):[];
    for(var i=0;i<grids.length;i++){
      var cards=directCards(grids[i]);
      for(var c=0;c<cards.length;c++){
        var cr=rect(cards[c]);
        if(!cr)continue;
        var links=cards[c].getElementsByTagName('a');
        var spans=cards[c].getElementsByTagName('span');
        for(var pass=0;pass<2;pass++){
          var nodes=pass==0?links:spans;
          for(var j=0;j<nodes.length;j++){
            var el=nodes[j];
            if(!hasClass(el,'tool'))continue;
            var br=rect(el);
            if(!br||br.w<1||br.h<1)continue;
            var txt=(el.innerText||el.textContent||'').replace(/\s+/g,' ').replace(/^\s+|\s+$/g,'');
            var textOverflow=(css(el,'textOverflow')||css(el,'text-overflow')||'').toLowerCase();
            var overflow=(css(el,'overflow')||css(el,'overflowX')||'').toLowerCase();
            if(textOverflow.indexOf('ellipsis')>=0)out.push(label+' button still uses ellipsis: '+txt);
            if(overflow.indexOf('hidden')>=0&&el.scrollWidth>el.clientWidth+2)out.push(label+' button content is clipped: '+txt);
            if(br.r>cr.r+3)out.push(label+' button escapes card right edge: '+txt);
            if(br.l<cr.l-3)out.push(label+' button escapes card left edge: '+txt);
          }
        }
      }
    }
  }
  function checkAuditTargetSelector(){
    var selectors=document.getElementsByClassName?document.getElementsByClassName('audit-target-selector'):[];
    if(!selectors||selectors.length<1){out.push('missing audit target selector');return;}
    var chips=selectors[0].getElementsByTagName('a');
    var active=0, routed=0;
    for(var i=0;i<chips.length;i++){
      var chip=chips[i];
      if(hasClass(chip,'audit-target-chip')){
        if(hasClass(chip,'active'))active++;
        var href=String(chip.getAttribute('href')||'');
        if(href.indexOf('kkyfb:browse-discipline-')==0)routed++;
        var label=(chip.innerText||chip.textContent||'').replace(/\s+/g,' ').replace(/^\s+|\s+$/g,'');
        var br=rect(chip), sr=rect(selectors[0]);
        if(br&&sr&&br.r>sr.r+3)out.push('audit target chip escapes selector: '+label);
      }
    }
    if(chips.length<2)out.push('audit target selector has fewer than two target chips');
    if(active!==1)out.push('audit target selector active chip count is '+active);
    if(routed<1)out.push('audit target selector has no browse-discipline route');
  }
  function checkAdminTradeControls(){
    var controls=document.getElementsByClassName?document.getElementsByClassName('admin-trade-control'):[];
    if(!controls||controls.length<1){out.push('missing admin trade control');return;}
    var control=controls[0];
    var selectors=control.getElementsByClassName?control.getElementsByClassName('admin-trade-selector'):[];
    var managers=control.getElementsByClassName?control.getElementsByClassName('admin-trade-management'):[];
    if(!selectors||selectors.length<1)out.push('admin trade selector is outside target control');
    if(!managers||managers.length<1)out.push('admin trade management is outside target control');
    if(selectors&&selectors.length){
      var links=selectors[0].getElementsByTagName('a'), active=0;
      for(var i=0;i<links.length;i++)if(hasClass(links[i],'active'))active++;
      if(active!==1)out.push('admin trade selector active target count is '+active);
    }
    if(managers&&managers.length){
      var actions=managers[0].getElementsByTagName('a');
      if(actions.length!==3)out.push('admin trade management action count is '+actions.length);
      var baseline=document.getElementsByClassName?document.getElementsByClassName('baseline-actions'):[];
      if(baseline&&baseline.length&&baseline[0].contains&&baseline[0].contains(managers[0]))out.push('trade management is still inside baseline RVT actions');
    }
    var grids=document.getElementsByClassName?document.getElementsByClassName('admin-standard-action-grid'):[];
    if(!grids||grids.length<1){out.push('missing admin standard action grid');return;}
    var rows=grids[0].getElementsByClassName?grids[0].getElementsByClassName('standard-action-row'):[];
    for(var r=0;r<rows.length;r++){
      var row=rows[r], all=row.getElementsByTagName('a'), buttons=[];
      for(var j=0;j<all.length;j++)if(hasClass(all[j],'standard-action-link'))buttons.push(all[j]);
      if(buttons.length<1)continue;
      var rr=rect(row), firstRect=rect(buttons[0]);
      if(!rr||!firstRect)continue;
      if(buttons.length===1&&firstRect.w<rr.w-4)out.push('single standard action does not fill its row');
      for(var k=1;k<buttons.length;k++){
        var br=rect(buttons[k]);
        if(!br)continue;
        if(Math.abs(br.w-firstRect.w)>4)out.push('standard action row has unequal button widths');
        if(Math.abs(br.h-firstRect.h)>4)out.push('standard action row has unequal button heights');
        if(Math.abs(br.t-firstRect.t)>3)out.push('standard action row buttons are vertically misaligned');
      }
    }
  }
  function checkAdminDetailedComponentOption(){
    var options=document.getElementsByClassName?document.getElementsByClassName('admin-check-option'):[];
    if(!options||options.length!==1){out.push('admin policy option count is '+(options?options.length:0));return;}
    var seenSystemDetail=false;
    for(var oi=0;oi<options.length;oi++){
      var option=options[oi], href=(option.getAttribute('href')||'').toLowerCase();
      if(href.indexOf('system-type-detail-components/')>=0)seenSystemDetail=true;
      if(href.indexOf('project-element-change-tracking/')>=0)out.push('removed global project tracking option is still visible');
      var leading=option.getElementsByClassName?option.getElementsByClassName('admin-check-leading'):[];
      var boxes=option.getElementsByClassName?option.getElementsByClassName('admin-check-box'):[];
      var copies=option.getElementsByClassName?option.getElementsByClassName('admin-check-copy'):[];
      var states=option.getElementsByClassName?option.getElementsByClassName('admin-check-state'):[];
      if(!leading.length||!boxes.length||!copies.length||!states.length){out.push('admin policy option cells are incomplete');continue;}
      var optionRect=rect(option), boxRect=rect(boxes[0]), copyRect=rect(copies[0]), stateRect=rect(states[0]);
      if(boxRect&&copyRect&&copyRect.l-boxRect.r<10)out.push('admin policy checkbox-to-copy gap is too small: '+(copyRect.l-boxRect.r));
      if(optionRect&&boxRect&&(boxRect.l<optionRect.l-2||boxRect.r>optionRect.r+2))out.push('admin policy checkbox escapes option');
      if(copyRect&&stateRect&&stateRect.l<copyRect.r-2)out.push('admin policy state overlaps copy');
      if(optionRect&&stateRect&&stateRect.r>optionRect.r+2)out.push('admin policy state escapes option');
      var headings=copies[0].getElementsByTagName('strong');
      var descriptions=copies[0].getElementsByTagName('span');
      if(!headings.length||!descriptions.length){out.push('admin policy title or description is missing');}
      else{
        var headingRect=rect(headings[0]), descriptionRect=rect(descriptions[0]);
        if(headingRect&&descriptionRect&&descriptionRect.t<headingRect.b+2)out.push('admin policy title and description are not vertically separated');
      }
      var statusText=(states[0].innerText||states[0].textContent||'').replace(/^\s+|\s+$/g,'');
      if(!statusText)out.push('admin policy state label is empty');
    }
    if(!seenSystemDetail)out.push('System Type detail comparison policy route is missing');
    var anchors=document.getElementsByTagName('a'), historyLinks=[], allHistoryLinks=[];
    for(var ai=0;ai<anchors.length;ai++){
      var actionHref=(anchors[ai].getAttribute('href')||'').toLowerCase();
      if(actionHref==='kkyfb:project-element-change-history')historyLinks.push(anchors[ai]);
      if(actionHref==='kkyfb:project-element-change-history-all')allHistoryLinks.push(anchors[ai]);
    }
    if(historyLinks.length!==1){out.push('project element change history action is missing or duplicated: '+historyLinks.length);}
    else{
      var historyLink=historyLinks[0];
      var label=(historyLink.innerText||historyLink.textContent||'').replace(/^\s+|\s+$/g,'');
      if(!label)out.push('project element change history action label is empty');
      if(historyLink.id!=='fbCurrentProjectHistoryTool')out.push('current-project history sidebar id is incorrect');
    }
    if(allHistoryLinks.length!==1){out.push('all-project element change history action is missing or duplicated: '+allHistoryLinks.length);}
    else{
      var allHistoryLink=allHistoryLinks[0];
      var allLabel=(allHistoryLink.innerText||allHistoryLink.textContent||'').replace(/^\s+|\s+$/g,'');
      if(!allLabel)out.push('all-project element change history action label is empty');
      if(allHistoryLink.id!=='fbAllProjectHistoryTool')out.push('all-project history sidebar id is incorrect');
    }
    var historyTitles=document.getElementsByClassName?document.getElementsByClassName('rail-history-title'):[];
    if(!historyTitles||historyTitles.length!==1)out.push('history sidebar group title count is '+(historyTitles?historyTitles.length:0));
    else if(historyLinks.length===1&&allHistoryLinks.length===1&&historyTitles[0].sourceIndex>=historyLinks[0].sourceIndex)out.push('history sidebar title is not before its actions');
  }
  if(hasBodyClass('fb-tab-admin')){
    checkAdminTradeControls();
    checkAdminDetailedComponentOption();
    checkTwoColumnActionCards('admin-standard-action-grid','admin standard action grid',false);
    checkActionCardButtonWrap('admin-standard-action-grid','admin standard action grid');
  }
  if(hasBodyClass('fb-tab-audit')){
    checkAuditTargetSelector();
    checkTwoColumnActionCards('audit-action-grid','audit action grid',true);
    checkActionCardButtonWrap('audit-action-grid','audit action grid');
  }
  if(hasBodyClass('fb-tab-permissions'))checkPermissionDiagnosticRows();
  return out.join('|')||'OK';
})()
";
		try
		{
			string value = Convert.ToString(browser.Document.InvokeScript("eval", new object[] { script }));
			if (!string.Equals(value, "OK", StringComparison.OrdinalIgnoreCase))
			{
				result.Failures.Add("Layout check failed: " + value);
			}
		}
		catch (Exception ex)
		{
			result.Failures.Add("Layout check threw " + ex.GetType().Name + ": " + ex.Message);
		}
	}

	private static void CheckOverflowTitleBehavior(WebBrowser browser, AuditResult result)
	{
		string validation = EvalScript(browser, @"
(function(){
  if(!window.KKYFBOverflowTitles||!window.refreshOverflowTitles)return 'FAIL overflow title service missing';
  var host=document.createElement('div');
  host.id='kkyfbOverflowAuditFixture';
  host.style.position='fixed';host.style.left='4px';host.style.bottom='4px';host.style.width='64px';host.style.height='24px';host.style.whiteSpace='nowrap';host.style.overflow='hidden';host.style.textOverflow='ellipsis';host.style.zIndex='-1';
  host.innerText='A deliberately long Family Browser value used to verify clipped text behavior';
  document.body.appendChild(host);
  try{
    window.refreshOverflowTitles(host);
    if(host.getAttribute('data-kkyfb-overflow-title')!=='1')return 'FAIL clipped text did not receive generated title';
    if((host.getAttribute('title')||'').indexOf('deliberately long Family Browser value')<0)return 'FAIL generated title lost full text';
    host.style.width='1100px';
    window.refreshOverflowTitles(host);
    if(host.getAttribute('data-kkyfb-overflow-title')==='1'||host.getAttribute('title'))return 'FAIL widened text kept stale generated title';
    host.style.width='32px';host.setAttribute('title','Authored help text');
    window.refreshOverflowTitles(host);
    if(host.getAttribute('title')!=='Authored help text'||host.getAttribute('data-kkyfb-overflow-title')==='1')return 'FAIL authored title was overwritten';
    return 'OK';
  }finally{
    if(host.parentNode)host.parentNode.removeChild(host);
  }
})()");
		if (!string.Equals(validation, "OK", StringComparison.Ordinal))
		{
			result.Failures.Add("Overflow title behavior check failed: " + validation);
		}
	}

	private static void CheckThemeBehavior(WebBrowser browser, Dictionary<string, string> options, AuditResult result)
	{
		const string expected = "light";
		if (!string.Equals(BodyAttribute(browser, "data-theme"), expected, StringComparison.OrdinalIgnoreCase))
		{
			result.Failures.Add("Dashboard must use the fixed product palette. Expected light, actual " + BodyAttribute(browser, "data-theme") + ".");
		}
		string bodyClass = browser.Document == null || browser.Document.Body == null ? string.Empty : browser.Document.Body.GetAttribute("className");
		if (bodyClass.IndexOf("theme-light", StringComparison.OrdinalIgnoreCase) < 0 || bodyClass.IndexOf("theme-dark", StringComparison.OrdinalIgnoreCase) >= 0)
		{
			result.Failures.Add("Dashboard fixed palette class is invalid: " + bodyClass);
		}
		HtmlElement toggle = browser.Document == null ? null : browser.Document.GetElementById("fbThemeToggle");
		if (toggle != null)
		{
			result.Failures.Add("Theme control must not be visible in the fixed-palette product UI.");
		}
		string visual = EvalString(browser, @"
(function(){
  function css(el,name){var s=window.getComputedStyle?window.getComputedStyle(el,null):el.currentStyle;return s?(s[name]||''):'';}
  var expected=document.body.getAttribute('data-theme')=='dark'?'rgb(11, 18, 32)':'rgb(245, 247, 251)';
  var bodyColor=css(document.body,'backgroundColor');
  var old={'rgb(29, 138, 98)':1,'rgb(22, 139, 96)':1,'rgb(16, 58, 48)':1,'rgb(18, 45, 39)':1,'rgb(32, 51, 44)':1,'rgb(31, 153, 109)':1,'rgb(37, 182, 127)':1,'rgb(232, 247, 240)':1};
  var selectors=['.top','.toolbar.fb-nav-10','.toolbar.fb-nav-10 a.tool.primary','.toolbar.fb-nav-10 a.tool.active','a.tool.primary','span.tool.primary','a.filter.active','.selector a.active','.command-apply','.inline-status-filter.active','.family-kind-toggle a.active','.settings-note','.admin-required','.admin-flow-number','.pill.action.primary','.btn.primary','.header','.run.primary','.empty-action'];var bad=[];
  for(var i=0;i<selectors.length;i++){var nodes=document.querySelectorAll?document.querySelectorAll(selectors[i]):[];for(var j=0;j<nodes.length&&j<24;j++){var bg=css(nodes[j],'backgroundColor');if(old[bg])bad.push(selectors[i]+'='+bg);}}
  return bodyColor+'|'+expected+'|'+bad.join(',');
})()");
		string[] visualParts = visual.Split(new[] { '|' }, 3);
		if (visualParts.Length < 3 || !string.Equals(visualParts[0], visualParts[1], StringComparison.OrdinalIgnoreCase))
		{
			result.Failures.Add("Theme body background did not resolve to the expected KKY palette: " + visual);
		}
		else if (!string.IsNullOrWhiteSpace(visualParts[2]))
		{
			result.Failures.Add("Legacy green brand styling remains on active controls: " + visualParts[2]);
		}
		string brand = EvalString(browser, @"
(function(){
  function rect(el){return el&&el.getBoundingClientRect?el.getBoundingClientRect():null;}
  function visible(el){if(!el)return false;var s=window.getComputedStyle?window.getComputedStyle(el,null):el.currentStyle;var r=rect(el);return !!(r&&r.width>0&&r.height>0&&(!s||s.display!='none')&&(!s||s.visibility!='hidden'));}
  var lock=document.querySelector?document.querySelector('.brand-lockup'):null;
  var mark=lock&&lock.querySelector?lock.querySelector('.brand-mark'):null;
  var title=lock&&lock.querySelector?lock.querySelector('.title'):null;
  var kicker=lock&&lock.querySelector?lock.querySelector('.kicker'):null;
  var rail=document.querySelector?document.querySelector('.toolbar'):null;
  if(!visible(lock))return 'missing-lockup';
  if(!visible(mark))return 'missing-logo';
  if(!visible(title)||!visible(kicker))return 'missing-copy';
  var lr=rect(lock),mr=rect(mark),tr=rect(title),kr=rect(kicker),rr=rect(rail);
  if(tr.width<140||tr.height<16||kr.width<80||kr.height<10)return 'collapsed-copy:'+tr.width+','+tr.height+','+kr.width+','+kr.height;
  if(rr&&lr.left<rr.right-1)return 'sidebar-overlap:'+lr.left+','+rr.right;
  if(tr.left<mr.right+6)return 'logo-copy-overlap:'+tr.left+','+mr.right;
  return 'OK';
})()");
		if (!string.Equals(brand, "OK", StringComparison.OrdinalIgnoreCase))
		{
			result.Failures.Add("Brand lockup layout is not usable: " + brand);
		}
		string contrast = EvalString(browser, @"
(function(){
  function css(el,name){var s=window.getComputedStyle?window.getComputedStyle(el,null):el.currentStyle;return s?(s[name]||''):'';}
  function rgb(value){var m=(value||'').match(/rgba?\s*\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)/i);return m?[parseInt(m[1],10),parseInt(m[2],10),parseInt(m[3],10)]:null;}
  function lum(c){function f(v){v=v/255;return v<=.03928?v/12.92:Math.pow((v+.055)/1.055,2.4);}return .2126*f(c[0])+.7152*f(c[1])+.0722*f(c[2]);}
  function ratio(a,b){var x=lum(a),y=lum(b);return (Math.max(x,y)+.05)/(Math.min(x,y)+.05);}
  function bg(el){while(el&&el.nodeType==1){var value=css(el,'backgroundColor');if(value&&value!='transparent'&&value!='rgba(0, 0, 0, 0)'){var c=rgb(value);if(c)return c;}el=el.parentNode;}return rgb(css(document.body,'backgroundColor'))||[255,255,255];}
  var selectors=['.brand-lockup .title','.brand-lockup .kicker','.home-hero-copy strong','.home-board-head strong','.home-discipline-table th','.home-discipline-table td','#browserSearch','.family-tablewrap','.section','.settings-note','.admin-flow-title','.admin-flow-value','.admin-guide','.admin-box .label','.settings-action-title','.settings-action-body','.tree-title','.tree-help','.tree-node','.family-list-panel th','.family-list-panel td','.pane-head','.statusbar','.detail-name','.detail-block-title','.detail-text'];
  var bad=[];
  for(var i=0;i<selectors.length;i++){
    var nodes=document.querySelectorAll?document.querySelectorAll(selectors[i]):[];
    for(var j=0;j<nodes.length&&j<8;j++){
      var el=nodes[j],r=el.getBoundingClientRect?el.getBoundingClientRect():null,text=(el.innerText||el.textContent||'').replace(/\s+/g,'');
      if(!text||!r||r.width<=0||r.height<=0)continue;
      var fg=rgb(css(el,'color')),back=bg(el);if(fg&&back&&ratio(fg,back)<3)bad.push(selectors[i]+'='+ratio(fg,back).toFixed(2));
    }
  }
  return bad.join(',');
})()");
		if (!string.IsNullOrWhiteSpace(contrast))
		{
			result.Failures.Add("Low-contrast themed content detected: " + contrast);
		}
	}

	private static string NormalizeThemeCode(string value)
	{
		return string.Equals((value ?? string.Empty).Trim(), "dark", StringComparison.OrdinalIgnoreCase) ? "dark" : "light";
	}

	private static string JsStringLiteral(string value)
	{
		return "'" + (value ?? string.Empty).Replace("\\", "\\\\").Replace("'", "\\'").Replace("\r", "\\r").Replace("\n", "\\n") + "'";
	}

	private static void CheckAuditTargetScrollPersistence(WebBrowser browser, AuditResult result)
	{
		string script = @"
(function(){
  var bodyClass=' '+((document.body&&document.body.className)||'')+' ';
  if(bodyClass.indexOf(' fb-tab-audit ')<0)return 'SKIP:not-audit';
  if(typeof captureDashboardUiStateJson!='function'||typeof restoreDashboardUiStateJson!='function')return 'FAIL:missing-state-functions';
  var pane=document.getElementById('auditPane');
  var center=document.getElementById('mainCenter');
  if(!pane||!center)return 'FAIL:missing-scroll-elements';
  var spacer=document.createElement('div');
  spacer.id='auditScrollAuditSpacer';
  spacer.style.height='1600px';
  spacer.style.width='1px';
  spacer.style.clear='both';
  pane.appendChild(spacer);
  var paneMax=Math.max(0,pane.scrollHeight-pane.clientHeight);
  var centerMax=Math.max(0,center.scrollHeight-center.clientHeight);
  var target=paneMax>=centerMax?pane:center;
  var targetName=target===pane?'workflow':'center';
  var max=Math.max(0,target.scrollHeight-target.clientHeight);
  if(max<120){pane.removeChild(spacer);return 'FAIL:no-scroll-surface';}
  target.scrollTop=Math.min(320,max);
  var expected=target.scrollTop;
  var state=captureDashboardUiStateJson();
  target.scrollTop=0;
  var restored=restoreDashboardUiStateJson(state);
  var actual=target.scrollTop;
  pane.removeChild(spacer);
  if(!restored||Math.abs(actual-expected)>2)return 'FAIL:'+targetName+':'+expected+':'+actual+':'+restored;
  return 'PASS:'+targetName+':'+expected+':'+actual;
})()
";
		try
		{
			string value = Convert.ToString(browser.Document.InvokeScript("eval", new object[] { script })) ?? string.Empty;
			if (!value.StartsWith("PASS:", StringComparison.OrdinalIgnoreCase) && !value.StartsWith("SKIP:", StringComparison.OrdinalIgnoreCase))
			{
				result.Failures.Add("Audit target scroll state did not round-trip: " + value);
			}
		}
		catch (Exception ex)
		{
			result.Failures.Add("Audit target scroll state check threw " + ex.GetType().Name + ": " + ex.Message);
		}
	}

	private static void CheckDebugDock(WebBrowser browser, AuditResult result)
	{
		string script = @"
(function(){
  function rect(el){if(!el||!el.getBoundingClientRect)return null;var r=el.getBoundingClientRect();return {l:r.left,t:r.top,r:r.right,b:r.bottom,w:r.right-r.left,h:r.bottom-r.top};}
  function first(cls){var x=document.getElementsByClassName?document.getElementsByClassName(cls):[];return x&&x.length?x[0]:null;}
  function css(el,name){try{var s=(window.getComputedStyle?window.getComputedStyle(el,null):el.currentStyle);return s?(s[name]||s.getPropertyValue&&s.getPropertyValue(name)||''):'';}catch(e){return '';}}
  var out=[];
  if(document.getElementById('fbDebugFab'))out.push('floating debug FAB still rendered');
  var panel=document.getElementById('fbDebug');
  if(!panel)return out.join('|')||'OK';
  if(typeof toggleDebug!='function')out.push('debug toggle function missing');
  var body=document.body;
  var bodyClass=' '+String((body&&body.className)||'')+' ';
  if(bodyClass.indexOf(' fb-debug-on ')<0 && typeof toggleDebug=='function')toggleDebug();
  var pr=rect(panel), tr=rect(first('toolbar'));
  var root=document.documentElement||document.body;
  var viewportH=(root&&root.clientHeight)||document.body.clientHeight||window.innerHeight||0;
  if(!pr||pr.w<1||pr.h<1){
    out.push('debug dock is not visible after toggle');
  }else{
    var display=String(css(panel,'display')||'').toLowerCase();
    if(display=='none')out.push('debug dock display remains none after toggle');
    if(viewportH>0 && Math.abs(pr.b-viewportH)>6)out.push('debug dock is not attached to viewport bottom: bottom='+pr.b+', viewport='+viewportH);
    if(viewportH>0 && pr.t<viewportH*0.45)out.push('debug dock is too high; expected bottom console: top='+pr.t+', viewport='+viewportH);
    if(tr&&pr.l<tr.r-2)out.push('debug dock overlaps left menu: panel.left='+pr.l+', toolbar.right='+tr.r);
    if(pr.h<260||pr.h>380)out.push('debug dock height is outside expected console range: '+pr.h);
  }
  if(typeof toggleDebug=='function')toggleDebug();
  return out.join('|')||'OK';
})()
";
		try
		{
			string value = Convert.ToString(browser.Document.InvokeScript("eval", new object[] { script }));
			if (!string.Equals(value, "OK", StringComparison.OrdinalIgnoreCase))
			{
				result.Failures.Add("Debug dock check failed: " + value);
			}
		}
		catch (Exception ex)
		{
			result.Failures.Add("Debug dock check threw " + ex.GetType().Name + ": " + ex.Message);
		}
	}

	private static void CheckProjectSubtitle(WebBrowser browser, Dictionary<string, string> options, AuditResult result)
	{
		string configuredProjectPath = GetOption(options, "projectPath", @"C:\KKY Audit\Local\UI_Audit_Local.rvt").Trim();
		string configuredCentralPath = GetOption(options, "centralPath", @"C:\KKY Audit\Central\UI_Audit_Central.rvt").Trim();
		bool placeholderProject = string.Equals(configuredProjectPath, "Unsaved project", StringComparison.OrdinalIgnoreCase) || string.Equals(configuredProjectPath, "저장되지 않은 프로젝트", StringComparison.Ordinal);
		string script;
		if (placeholderProject)
		{
			string expectedProjectText = string.Equals(NormalizeLanguageCode(GetOption(options, "languageCode", "ko")), "en", StringComparison.OrdinalIgnoreCase) ? "Unsaved project" : "저장되지 않은 프로젝트";
			script = @"
(function(){
  var rows=document.getElementsByClassName('project-file-row'),out=[];
  var local=null,central=null;
  for(var i=0;i<rows.length;i++){var c=' '+String(rows[i].className||'')+' ';if(c.indexOf(' local ')>=0)local=rows[i];if(c.indexOf(' central ')>=0)central=rows[i];}
  if(!local)out.push('missing local project status token');
  var text=local?String(local.innerText||local.textContent||'').replace(/\s+/g,' '):'';
  if(local&&text.indexOf(" + JsStringLiteral(expectedProjectText) + @")<0)out.push('unsaved project text mismatch: '+text);
  if(central)out.push('unsaved project unexpectedly shows central token');
  return out.join('|')||'OK';
})()
";
		}
		else
		{
			string expectedLocalFileName = Path.GetFileName(configuredProjectPath);
			string expectedCentralFileName = Path.GetFileName(configuredCentralPath);
			script = @"
(function(){
  function first(cls){var x=document.getElementsByClassName(cls);return x&&x.length?x[0]:null;}
  function text(el){return el?(el.innerText||el.textContent||''):'';}
  function rect(el){if(!el||!el.getBoundingClientRect)return null;var r=el.getBoundingClientRect();return {l:r.left,t:r.top,r:r.right,b:r.bottom,w:r.right-r.left,h:r.bottom-r.top};}
  var out=[];
  var subtitle=first('project-subtitle');
  var stack=first('project-file-stack');
  var context=first('project-context');
  if(!subtitle){out.push('missing project subtitle');return out.join('|');}
  if(!stack)out.push('missing stacked project file layout');
  if(!context)out.push('missing right-side project context');
  if(first('project-title'))out.push('duplicate project title token is still rendered');
  var local=first('local');
  var central=first('central');
  if(!local)out.push('missing local project file token');
  if(" + JsStringLiteral(expectedCentralFileName) + @"&&!central)out.push('missing central project file token');
  if(local&&text(local).indexOf(" + JsStringLiteral(expectedLocalFileName) + @")<0)out.push('local file name missing: '+text(local));
  if(central&&" + JsStringLiteral(expectedCentralFileName) + @"&&text(central).indexOf(" + JsStringLiteral(expectedCentralFileName) + @")<0)out.push('central file name missing: '+text(central));
  if(local&&String(local.getAttribute('title')||'').indexOf(" + JsStringLiteral(configuredProjectPath) + @")<0)out.push('local title path missing');
  if(central&&" + JsStringLiteral(configuredCentralPath) + @"&&String(central.getAttribute('title')||'').indexOf(" + JsStringLiteral(configuredCentralPath) + @")<0)out.push('central title path missing');
  var lr=rect(local), cr=rect(central);
  if(lr&&cr&&Math.abs(lr.t-cr.t)<2)out.push('local and central project files are still on one line');
  return out.join('|')||'OK';
})()
";
		}
		try
		{
			string value = Convert.ToString(browser.Document.InvokeScript("eval", new object[] { script }));
			if (!string.Equals(value, "OK", StringComparison.OrdinalIgnoreCase))
			{
				result.Failures.Add("Project subtitle check failed: " + value);
			}
		}
		catch (Exception ex)
		{
			result.Failures.Add("Project subtitle check threw " + ex.GetType().Name + ": " + ex.Message);
		}
	}

	private static void CheckLanguagePurity(WebBrowser browser, Dictionary<string, string> options, AuditResult result)
	{
		string language = NormalizeLanguageCode(GetOption(options, "languageCode", "ko"));
		List<string> snippets = CollectVisibleTextSnippets(browser);
		if (snippets.Count == 0)
		{
			result.Failures.Add("Language check failed: no visible text snippets were collected.");
			return;
		}
		string windowTitle = EvalString(browser, "(function(){return document.body?String(document.body.getAttribute('data-window-title')||''):'';})()");
		if (string.Equals(language, "en", StringComparison.OrdinalIgnoreCase))
		{
			if (string.IsNullOrWhiteSpace(windowTitle) || HangulRegex.IsMatch(windowTitle))
			{
				result.Failures.Add("Language check failed: English window title is missing or contains Korean text: " + windowTitle);
			}
			List<string> leaks = snippets
				.Where(snippet => HangulRegex.IsMatch(snippet) && !IsAllowedEnglishModeHangulSnippet(snippet))
				.Select(CompactLanguageSnippet)
				.Distinct(StringComparer.Ordinal)
				.Take(12)
				.ToList();
			if (leaks.Count > 0)
			{
				result.Failures.Add("Language check failed: English mode contains Korean visible text: " + string.Join(" | ", leaks.ToArray()));
			}
			return;
		}

		if (string.IsNullOrWhiteSpace(windowTitle) || windowTitle.IndexOf("패밀리 브라우저", StringComparison.Ordinal) < 0)
		{
			result.Failures.Add("Language check failed: Korean window title is missing: " + windowTitle);
		}
		string allText = string.Join(" ", snippets.ToArray());
		string[] requiredKoreanTokens = new[] { "홈", "패밀리", "시스템 타입", "요청", "표준", "새로고침" };
		List<string> missingTokens = requiredKoreanTokens.Where(token => allText.IndexOf(token, StringComparison.Ordinal) < 0).ToList();
		if (missingTokens.Count > 0)
		{
			result.Failures.Add("Language check failed: Korean mode is missing required Korean UI tokens: " + string.Join(", ", missingTokens.ToArray()));
		}

		List<string> englishLeaks = snippets
			.Where(ContainsDisallowedEnglishUiInKoreanMode)
			.Select(CompactLanguageSnippet)
			.Distinct(StringComparer.Ordinal)
			.Take(12)
			.ToList();
		if (englishLeaks.Count > 0)
		{
			result.Failures.Add("Language check failed: Korean mode contains untranslated English UI text: " + string.Join(" | ", englishLeaks.ToArray()));
		}
	}

	private static void CheckAdminOffFileGuardUi(WebBrowser browser, Dictionary<string, string> options, AuditResult result)
	{
		if (!GetBool(options, "expectFamilyLoadBlocked", false))
		{
			return;
		}
		if (!string.Equals(GetOption(options, "activeTab", string.Empty), "families", StringComparison.OrdinalIgnoreCase))
		{
			result.Failures.Add("Admin OFF File Guard UI audit must run on the Families tab.");
			return;
		}

		string state = EvalString(browser, @"
(function(){
  var pane=document.getElementById('familiesPane');
  if(!pane)return 'missing-pane';
  var links=pane.getElementsByTagName('a'),spans=pane.getElementsByTagName('span');
  var activeApply=0,disabledApply=0,label='';
  for(var i=0;i<links.length;i++){
    if(links[i].id==='checkedFamilyApply'||(' '+String(links[i].className||'')+' ').indexOf(' command-apply ')>=0){
      activeApply++;
      label=String(links[i].innerText||links[i].textContent||'');
    }
  }
  for(var j=0;j<spans.length;j++){
    var cls=' '+String(spans[j].className||'')+' ';
    if(cls.indexOf(' command-apply ')>=0&&cls.indexOf(' disabled ')>=0){
      disabledApply++;
      label=String(spans[j].innerText||spans[j].textContent||'');
    }
  }
	  var allowed=(typeof window.canLoadFamilies==='undefined')?'undefined':String(window.canLoadFamilies);
	  var modeSwitch=document.getElementById('adminModeSwitch');
	  var mode=modeSwitch?String(modeSwitch.getAttribute('data-admin-mode')||''):'missing';
	  var modeLinks=modeSwitch?modeSwitch.getElementsByTagName('a'):[];
	  var onHref='',offHref='',onActive='false',offActive='false';
	  for(var k=0;k<modeLinks.length;k++){
	    var href=String(modeLinks[k].getAttribute('href')||'');
	    var active=(' '+String(modeLinks[k].className||'')+' ').indexOf(' active ')>=0?'true':'false';
	    if(href.indexOf('admin-mode-on')>=0){onHref=href;onActive=active;}
	    if(href.indexOf('admin-mode-off')>=0){offHref=href;offActive=active;}
	  }
	  return allowed+'|'+activeApply+'|'+disabledApply+'|'+label.replace(/\|/g,'/')+'|'+mode+'|'+onHref+'|'+offHref+'|'+onActive+'|'+offActive;
	})()");
		if (string.IsNullOrWhiteSpace(state) || state.StartsWith("eval-error:", StringComparison.OrdinalIgnoreCase) || string.Equals(state, "missing-pane", StringComparison.OrdinalIgnoreCase))
		{
			result.Failures.Add("Admin OFF File Guard UI audit could not inspect the Families pane: " + state);
			return;
		}

		string[] parts = state.Split(new[] { '|' }, 9);
		int activeApply;
		int disabledApply;
		if (parts.Length != 9 || !int.TryParse(parts[1], out activeApply) || !int.TryParse(parts[2], out disabledApply))
		{
			result.Failures.Add("Admin OFF File Guard UI audit returned invalid state: " + state);
			return;
		}
		if (!string.Equals(parts[0], "false", StringComparison.OrdinalIgnoreCase))
		{
			result.Failures.Add("Admin OFF File Guard UI audit: canLoadFamilies was not false: " + parts[0]);
		}
		if (activeApply != 0 || disabledApply != 1)
		{
			result.Failures.Add("Admin OFF File Guard UI audit: expected one disabled load control and no active load link, got active=" + activeApply + ", disabled=" + disabledApply + ".");
		}
		if (parts[3].IndexOf("선택 항목 로드", StringComparison.Ordinal) < 0 && parts[3].IndexOf("Load Selected Items", StringComparison.OrdinalIgnoreCase) < 0)
		{
			result.Failures.Add("Admin OFF File Guard UI audit: disabled load control has an unexpected label: " + parts[3]);
		}
		if (!string.Equals(parts[4], "off", StringComparison.OrdinalIgnoreCase))
		{
			result.Failures.Add("Admin OFF File Guard UI audit: explicit Admin switch does not report OFF: " + parts[4]);
		}
		if (parts[5].IndexOf("admin-mode-on", StringComparison.OrdinalIgnoreCase) < 0 || parts[6].IndexOf("admin-mode-off", StringComparison.OrdinalIgnoreCase) < 0)
		{
			result.Failures.Add("Admin OFF File Guard UI audit: explicit ON/OFF routes are incomplete.");
		}
		if (!string.Equals(parts[7], "false", StringComparison.OrdinalIgnoreCase) || !string.Equals(parts[8], "true", StringComparison.OrdinalIgnoreCase))
		{
			result.Failures.Add("Admin OFF File Guard UI audit: OFF segment is not the only active choice.");
		}
	}

	private static void CheckStandardSetupEmptyState(WebBrowser browser, Dictionary<string, string> options, AuditResult result)
	{
		string activeTab = GetOption(options, "activeTab", string.Empty);
		bool browserTab = string.Equals(activeTab, "families", StringComparison.OrdinalIgnoreCase) ||
			string.Equals(activeTab, "systems", StringComparison.OrdinalIgnoreCase);
		bool standardListRegistered = GetBool(options, "standardListRegistered", true);
		if (!browserTab || standardListRegistered)
		{
			return;
		}

		string paneId = string.Equals(activeTab, "systems", StringComparison.OrdinalIgnoreCase) ? "systemsPane" : "familiesPane";
		string script = @"
(function(){
  var pane=document.getElementById('" + paneId + @"');
  if(!pane)return 'missing-pane';
  var links=pane.getElementsByTagName('a'),listCount=0,rvtCount=0;
  for(var i=0;i<links.length;i++){
    var href=String(links[i].getAttribute('href')||links[i].href||'').toLowerCase();
    if(href.indexOf('open-standard-list-registration')>=0)listCount++;
    if(href.indexOf('open-standard-registration')>=0)rvtCount++;
  }
  var text=String(pane.innerText||pane.textContent||'').replace(/\s+/g,' ');
  return listCount+'|'+rvtCount+'|'+text;
})()
";
		string state = EvalString(browser, script);
		if (string.IsNullOrWhiteSpace(state) || state.StartsWith("eval-error:", StringComparison.OrdinalIgnoreCase) || string.Equals(state, "missing-pane", StringComparison.OrdinalIgnoreCase))
		{
			result.Failures.Add("Standard setup empty-state check failed to inspect " + paneId + ": " + state);
			return;
		}

		string[] parts = state.Split(new[] { '|' }, 3);
		int listCount;
		int rvtCount;
		if (parts.Length != 3 || !int.TryParse(parts[0], out listCount) || !int.TryParse(parts[1], out rvtCount))
		{
			result.Failures.Add("Standard setup empty-state check returned invalid state: " + state);
			return;
		}

		bool standardRvtRegistered = GetBool(options, "standardRvtRegistered", true);
		bool adminMode = GetBool(options, "adminMode", true);
		string language = NormalizeLanguageCode(GetOption(options, "languageCode", "ko"));
		string expectedText;
		if (standardRvtRegistered)
		{
			expectedText = string.Equals(language, "en", StringComparison.OrdinalIgnoreCase)
				? "Connect the approved list for the registered standard RVT"
				: "등록된 표준 RVT의 표준 목록을 연결해주세요";
			int expectedListCount = adminMode ? 1 : 0;
			if (listCount != expectedListCount)
			{
				result.Failures.Add("Registered RVT / missing list state has wrong standard-list CTA count in " + activeTab + ": " + listCount + ".");
			}
			if (rvtCount != 0)
			{
				result.Failures.Add("Registered RVT / missing list state incorrectly shows the standard-RVT registration CTA in " + activeTab + ".");
			}
		}
		else
		{
			expectedText = string.Equals(language, "en", StringComparison.OrdinalIgnoreCase)
				? "Register the standard RVT first"
				: "표준 RVT를 먼저 등록해주세요";
			int expectedRvtCount = adminMode ? 1 : 0;
			if (rvtCount != expectedRvtCount)
			{
				result.Failures.Add("Missing RVT state has wrong standard-RVT registration CTA count in " + activeTab + ": " + rvtCount + ".");
			}
			if (listCount != 0)
			{
				result.Failures.Add("Missing RVT state incorrectly shows the standard-list registration CTA in " + activeTab + ".");
			}
		}

		if (parts[2].IndexOf(expectedText, StringComparison.Ordinal) < 0)
		{
			result.Failures.Add("Standard setup empty-state message mismatch in " + activeTab + ". Expected: " + expectedText);
		}
	}

	private static void CheckStandardRevisionState(WebBrowser browser, Dictionary<string, string> options, AuditResult result)
	{
		bool changed = GetBool(options, "standardRvtChanged", false);
		bool unavailable = GetBool(options, "standardRvtUnavailable", false);
		if (!changed && !unavailable)
		{
			return;
		}
		string activeTab = GetOption(options, "activeTab", string.Empty);
		string paneId = string.Equals(activeTab, "systems", StringComparison.OrdinalIgnoreCase) ? "systemsPane" : "familiesPane";
		string state = EvalString(browser, @"
(function(){
  var body=String(document.body.innerText||document.body.textContent||'').replace(/\s+/g,' ');
  var pane=document.getElementById('" + paneId + @"'),rows=0;
  if(pane){
    var trs=pane.getElementsByTagName('tr');
    for(var i=0;i<trs.length;i++)if((' '+String(trs[i].className||'')+' ').indexOf(' data ')>=0)rows++;
  }
  var boards=document.getElementsByTagName('div'),badBoards=0;
  for(var j=0;j<boards.length;j++){
    var cls=' '+String(boards[j].className||'')+' ';
    if(cls.indexOf(' standard-revision-board ')>=0&&cls.indexOf(' bad ')>=0)badBoards++;
  }
  return rows+'|'+badBoards+'|'+body;
})()");
		if (string.IsNullOrWhiteSpace(state) || state.StartsWith("eval-error:", StringComparison.OrdinalIgnoreCase))
		{
			result.Failures.Add("Standard revision state audit could not inspect the dashboard: " + state);
			return;
		}
		string[] parts = state.Split(new[] { '|' }, 3);
		int rowCount;
		int badBoardCount;
		if (parts.Length != 3 || !int.TryParse(parts[0], out rowCount) || !int.TryParse(parts[1], out badBoardCount))
		{
			result.Failures.Add("Standard revision state audit returned invalid DOM state: " + state);
			return;
		}
		string language = NormalizeLanguageCode(GetOption(options, "languageCode", "ko"));
		string expected = changed
			? (string.Equals(language, "en", StringComparison.OrdinalIgnoreCase) ? "changed - rescan required" : "변경됨 - 재스캔 필요")
			: (string.Equals(language, "en", StringComparison.OrdinalIgnoreCase) ? "source unavailable" : "원본 연결 불가");
		if (parts[2].IndexOf(expected, StringComparison.OrdinalIgnoreCase) < 0)
		{
			result.Failures.Add("Standard revision state audit is missing the expected status text: " + expected);
		}
		if ((string.Equals(activeTab, "families", StringComparison.OrdinalIgnoreCase) || string.Equals(activeTab, "systems", StringComparison.OrdinalIgnoreCase)) && rowCount != 0)
		{
			result.Failures.Add("Standard revision state audit: stale browser rows remained visible after source verification was blocked.");
		}
		if (string.Equals(activeTab, "home", StringComparison.OrdinalIgnoreCase) && badBoardCount < 1)
		{
			result.Failures.Add("Standard revision state audit: Home does not show the blocked Standard RVT source board.");
		}
	}

	private static void CheckManagedFolderRecovery(WebBrowser browser, Dictionary<string, string> options, AuditResult result)
	{
		bool expected = GetBool(options, "managedFolderUnavailable", false);
		string state = EvalString(browser, @"
(function(){
  var banner=document.getElementById('managedFolderRecovery');
  if(!banner)return 'missing';
  var links=banner.getElementsByTagName('a'),retry=0,setup=0;
  for(var i=0;i<links.length;i++){
    var href=String(links[i].getAttribute('href')||links[i].href||'').toLowerCase();
    if(href.indexOf('managed-folder-retry')>=0)retry++;
    if(href.indexOf('managed-folder-test-setup')>=0)setup++;
  }
  var text=String(banner.innerText||banner.textContent||'').replace(/\s+/g,' ');
  return retry+'|'+setup+'|'+text;
})()");
		if (!expected)
		{
			if (!string.Equals(state, "missing", StringComparison.OrdinalIgnoreCase))
			{
				result.Failures.Add("Managed-folder recovery banner is visible even though the scenario has a usable management folder.");
			}
			return;
		}
		if (string.IsNullOrWhiteSpace(state) || state.StartsWith("eval-error:", StringComparison.OrdinalIgnoreCase) || string.Equals(state, "missing", StringComparison.OrdinalIgnoreCase))
		{
			result.Failures.Add("Managed-folder recovery banner is missing when the homepage management folder is unavailable.");
			return;
		}
		string[] parts = state.Split(new[] { '|' }, 3);
		int retryCount;
		int setupCount;
		if (parts.Length != 3 || !int.TryParse(parts[0], out retryCount) || !int.TryParse(parts[1], out setupCount) || retryCount != 1 || setupCount != 1)
		{
			result.Failures.Add("Managed-folder recovery banner must expose exactly one homepage retry and one TEST folder setup action: " + state);
			return;
		}
		string language = NormalizeLanguageCode(GetOption(options, "languageCode", "ko"));
		string expectedSetup = string.Equals(language, "en", StringComparison.OrdinalIgnoreCase) ? "Set Up TEST Management Folder" : "TEST 관리폴더 설정";
		string expectedWarning = string.Equals(language, "en", StringComparison.OrdinalIgnoreCase) ? "Do not manually edit" : "수동으로 수정";
		if (parts[2].IndexOf(expectedSetup, StringComparison.Ordinal) < 0 || parts[2].IndexOf(expectedWarning, StringComparison.Ordinal) < 0)
		{
			result.Failures.Add("Managed-folder recovery banner is missing its TEST setup label or generated-file warning.");
		}
	}

	private static void CheckManagedFolderTransition(WebBrowser browser, Dictionary<string, string> options, AuditResult result)
	{
		bool testOverride = GetBool(options, "managedFolderTestOverride", false);
		bool homepageAvailable = GetBool(options, "homepageManagedFolderAvailable", false);
		bool adminMode = GetBool(options, "adminMode", true);
		string state = EvalString(browser, @"
(function(){
  var banner=document.getElementById('managedFolderOverride');
  if(!banner)return 'missing';
  var links=banner.getElementsByTagName('a'),retry=0,switchOnly=0,migrate=0;
  for(var i=0;i<links.length;i++){
    var href=String(links[i].getAttribute('href')||links[i].href||'').toLowerCase();
    if(href.indexOf('managed-folder-retry')>=0)retry++;
    if(href.indexOf('managed-folder-switch-homepage')>=0)switchOnly++;
    if(href.indexOf('managed-folder-migrate-homepage')>=0)migrate++;
  }
  var text=String(banner.innerText||banner.textContent||'').replace(/\s+/g,' ');
  return retry+'|'+switchOnly+'|'+migrate+'|'+text;
})()");
		if (!testOverride)
		{
			if (!string.Equals(state, "missing", StringComparison.OrdinalIgnoreCase))
			{
				result.Failures.Add("Managed-folder transition panel is visible without a TEST override.");
			}
			return;
		}
		if (string.IsNullOrWhiteSpace(state) || state.StartsWith("eval-error:", StringComparison.OrdinalIgnoreCase) || string.Equals(state, "missing", StringComparison.OrdinalIgnoreCase))
		{
			result.Failures.Add("Managed-folder transition panel is missing while a TEST override is active.");
			return;
		}
		string[] parts = state.Split(new[] { '|' }, 4);
		int retryCount;
		int switchCount;
		int migrateCount;
		if (parts.Length != 4 || !int.TryParse(parts[0], out retryCount) || !int.TryParse(parts[1], out switchCount) || !int.TryParse(parts[2], out migrateCount))
		{
			result.Failures.Add("Managed-folder transition panel state could not be parsed: " + state);
			return;
		}
		int expectedSwitch = homepageAvailable ? 1 : 0;
		int expectedMigrate = homepageAvailable && adminMode ? 1 : 0;
		if (retryCount != 1 || switchCount != expectedSwitch || migrateCount != expectedMigrate)
		{
			result.Failures.Add("Managed-folder transition actions do not match homepage availability/admin state: " + state);
		}
		string language = NormalizeLanguageCode(GetOption(options, "languageCode", "ko"));
		string expectedActive = string.Equals(language, "en", StringComparison.OrdinalIgnoreCase) ? "TEST MANAGEMENT FOLDER ACTIVE" : "TEST 관리폴더 사용 중";
		string expectedSourceSafety = string.Equals(language, "en", StringComparison.OrdinalIgnoreCase) ? "never deletes the TEST source" : "TEST 원본은 삭제하지";
		if (parts[3].IndexOf(expectedActive, StringComparison.Ordinal) < 0 || parts[3].IndexOf(expectedSourceSafety, StringComparison.Ordinal) < 0)
		{
			result.Failures.Add("Managed-folder transition panel is missing its TEST state or source-retention warning.");
		}
		if (homepageAvailable)
		{
			string expectedTarget = string.Equals(language, "en", StringComparison.OrdinalIgnoreCase) ? "Homepage destination" : "홈페이지 대상";
			if (parts[3].IndexOf(expectedTarget, StringComparison.Ordinal) < 0)
			{
				result.Failures.Add("Managed-folder transition panel does not show the reachable homepage destination.");
			}
		}
	}

	private static void CheckPendingTrackingQueue(WebBrowser browser, Dictionary<string, string> options, AuditResult result)
	{
		int expected = ParseInt(GetOption(options, "trackingPendingCount", "0"), 0);
		string state = EvalString(browser, @"
(function(){
  var pill=document.getElementById('pendingTrackingPill');
  var board=document.getElementById('pendingTrackingQueue');
  var retry=0;
  if(board){
    var links=board.getElementsByTagName('a');
    for(var i=0;i<links.length;i++){
      var href=links[i].getAttribute('href')||'';
      if(href.indexOf('homepage-security-refresh')>=0)retry++;
    }
  }
  return (pill?'1':'0')+'|'+(board?'1':'0')+'|'+retry+'|'+(pill?(pill.innerText||pill.textContent||''):'')+'|'+(board?(board.innerText||board.textContent||''):'');
})()");
		string[] parts = (state ?? string.Empty).Split('|');
		bool pillVisible = parts.Length > 0 && parts[0] == "1";
		bool boardVisible = parts.Length > 1 && parts[1] == "1";
		int retryCount = parts.Length > 2 ? ParseInt(parts[2], -1) : -1;
		if (expected <= 0)
		{
			if (pillVisible || boardVisible)
			{
				result.Failures.Add("Pending tracking warning is visible without queued records: " + state);
			}
			return;
		}
		if (!pillVisible || !boardVisible || retryCount != 1)
		{
			result.Failures.Add("Pending tracking queue must show one header pill, one Home board, and one retry action: " + state);
			return;
		}
		string expectedCount = expected.ToString(CultureInfo.InvariantCulture);
		if ((state ?? string.Empty).IndexOf(expectedCount, StringComparison.Ordinal) < 0)
		{
			result.Failures.Add("Pending tracking queue does not show its record count: " + state);
		}
	}

	private static void CheckProjectCatalogState(WebBrowser browser, Dictionary<string, string> options, AuditResult result)
	{
		bool baselineMissing = GetBool(options, "projectCatalogBaselineMissing", false);
		bool changed = GetBool(options, "projectCatalogChanged", false);
		bool untracked = GetBool(options, "projectCatalogUntracked", false);
		if (!baselineMissing && !changed && !untracked)
		{
			return;
		}
		string state = EvalString(browser, @"
(function(){
  var board=null,all=document.getElementsByTagName('*');
  for(var i=0;i<all.length;i++){
    var cls=' '+(all[i].className||'')+' ';
    if(cls.indexOf(' project-catalog-board ')>=0){board=all[i];break;}
  }
  var boardText=board?String(board.innerText||board.textContent||'').replace(/\s+/g,' '):'';
  var pillText='';
  var spans=document.getElementsByTagName('span');
  for(var j=0;j<spans.length;j++){
    var text=String(spans[j].innerText||spans[j].textContent||'');
    if(text.indexOf('Project Catalog:')>=0||text.indexOf('프로젝트 카탈로그:')>=0){pillText=text;break;}
  }
  return (board?'1':'0')+'|'+pillText+'|'+boardText;
})()");
		string language = NormalizeLanguageCode(GetOption(options, "languageCode", "ko"));
		if (string.IsNullOrWhiteSpace(state) || state.StartsWith("eval-error:", StringComparison.OrdinalIgnoreCase) || !state.StartsWith("1|", StringComparison.Ordinal))
		{
			result.Failures.Add("Project Catalog state board or header pill is missing: " + state);
			return;
		}
		string expectedStatus = baselineMissing
			? (language == "en" ? "baseline required" : "기준선 필요")
			: (language == "en" ? "differences detected" : "차이 감지");
		if (state.IndexOf(expectedStatus, StringComparison.OrdinalIgnoreCase) < 0)
		{
			result.Failures.Add("Project Catalog state does not show the contracted status '" + expectedStatus + "': " + state);
		}
		if (changed && state.IndexOf("AUDIT_EXTERNAL_FAMILY", StringComparison.Ordinal) < 0)
		{
			result.Failures.Add("Changed Project Catalog does not render its item-level difference: " + state);
		}
		if (untracked)
		{
			string expectedSource = language == "en" ? "External / untracked" : "외부 / 미추적";
			if (state.IndexOf(expectedSource, StringComparison.OrdinalIgnoreCase) < 0)
			{
				result.Failures.Add("Untracked Project Catalog difference is not labelled as external/untracked: " + state);
			}
		}
	}

	private static string NormalizeLanguageCode(string language)
	{
		return string.Equals(language, "en", StringComparison.OrdinalIgnoreCase) ? "en" : "ko";
	}

	private static void CheckPendingCommitState(WebBrowser browser, Dictionary<string, string> options, AuditResult result)
	{
		if (!GetBool(options, "includePendingRows", false))
		{
			return;
		}
		string activeTab = GetOption(options, "activeTab", string.Empty);
		bool familyTab = string.Equals(activeTab, "families", StringComparison.OrdinalIgnoreCase);
		bool systemTab = string.Equals(activeTab, "systems", StringComparison.OrdinalIgnoreCase);
		if (!familyTab && !systemTab)
		{
			result.Failures.Add("Pending commit scenario must use the Families or System Types tab.");
			return;
		}

		string expectedRaw = familyTab ? "FamilyPendingSaveOrSync" : "SystemPendingSaveOrSync";
		string script = @"
(function(){
  var tab=window.currentTab||'';
  var store=window.KKYFB&&window.KKYFB._stores?window.KKYFB._stores[tab]:null;
  if(!store||!store.rows)return 'missing-store';
  var expected='" + expectedRaw + @"',match=null,count=0;
  for(var i=0;i<store.rows.length;i++){
    var raw=store.rows[i]&&store.rows[i].attrs?String(store.rows[i].attrs['data-raw']||''):'';
    if(raw==expected){match=store.rows[i];count++;}
  }
  if(!match)return 'missing-row';
  try{
    window.currentFilter='All';window.currentDiscipline='All';window.currentTreeDiscipline='All';window.currentTreeGroup='';window.currentTreeCategory='';
    window.currentSystemTreeDiscipline='All';window.currentSystemTreeCategory='';window.advStatus='All';window.advGroup='All';window.advCategory='';window.advMismatchOnly=false;
    if(window.filterRows)window.filterRows('pending-audit');
  }catch(ignoreFilter){}
  var table=document.getElementById(tab+'Table'),dom=null,disabled='missing';
  if(table){
    var rows=table.getElementsByTagName('tr');
    for(var r=0;r<rows.length;r++)if(String(rows[r].getAttribute('data-raw')||'')==expected){dom=rows[r];break;}
  }
  if(dom){
    var inputs=dom.getElementsByTagName('input');
    disabled=(inputs.length&&inputs[0].disabled)?'1':'0';
    try{if(window.selectRow)window.selectRow(dom,false);}catch(ignoreSelect){}
  }
  var detailStatus=document.getElementById('detailStatus');
  var detailNotes=document.getElementById('detailNotes');
  return [count,match.selectable?'1':'0',String(match.attrs['data-status']||''),String(match.attrs['data-action']||''),disabled,dom?'1':'0',detailStatus?String(detailStatus.innerText||detailStatus.textContent||''):'',detailNotes?String(detailNotes.innerText||detailNotes.textContent||''):''].join('\u001f');
})()
";
		string state = EvalString(browser, script);
		if (string.Equals(state, "missing-store", StringComparison.OrdinalIgnoreCase) || string.Equals(state, "missing-row", StringComparison.OrdinalIgnoreCase) || state.StartsWith("eval-error:", StringComparison.OrdinalIgnoreCase))
		{
			result.Failures.Add("Pending save/sync state check failed: " + state);
			return;
		}
		string[] parts = state.Split(new[] { '\u001f' });
		if (parts.Length < 8)
		{
			result.Failures.Add("Pending save/sync state returned invalid data: " + state);
			return;
		}
		if (parts[0] != "1")
		{
			result.Failures.Add("Pending save/sync state expected exactly one row but found " + parts[0] + ".");
		}
		if (parts[1] != "0" || parts[4] != "1")
		{
			result.Failures.Add("Pending save/sync row is still selectable before commit.");
		}
		if (parts[5] != "1")
		{
			result.Failures.Add("Pending save/sync row was not rendered in the active list.");
		}
		string language = NormalizeLanguageCode(GetOption(options, "languageCode", "ko"));
		string expectedStatus = familyTab
			? (language == "en" ? "Loaded - Save/Sync Pending" : "로드됨 · 저장/동기화 대기")
			: (language == "en" ? "Applied - Save/Sync Pending" : "적용됨 · 저장/동기화 대기");
		string expectedAction = language == "en" ? "Save or synchronize to confirm" : "저장 또는 동기화 후 확정";
		if (parts[2].IndexOf(expectedStatus, StringComparison.Ordinal) < 0 || parts[6].IndexOf(expectedStatus, StringComparison.Ordinal) < 0)
		{
			result.Failures.Add("Pending save/sync status label mismatch. Expected: " + expectedStatus + ".");
		}
		if (parts[3].IndexOf(expectedAction, StringComparison.Ordinal) < 0)
		{
			result.Failures.Add("Pending save/sync action label mismatch. Expected: " + expectedAction + ".");
		}
		string expectedDetailHeading = language == "en" ? "Save/Sync Pending" : "저장/동기화 대기";
		if (parts[7].IndexOf(expectedDetailHeading, StringComparison.Ordinal) < 0)
		{
			result.Failures.Add("Pending save/sync detail card is missing its warning heading.");
		}
	}

	private static void CheckNestedFamilyDifferenceState(WebBrowser browser, Dictionary<string, string> options, AuditResult result)
	{
		if (!string.Equals(GetOption(options, "activeTab", string.Empty), "families", StringComparison.OrdinalIgnoreCase) ||
			!GetBool(options, "includeRows", true) ||
			!GetBool(options, "standardRvtRegistered", true) ||
			!GetBool(options, "standardListRegistered", true))
		{
			return;
		}

		string state = EvalString(browser, @"
(function(){
  var out=[];
  var store=window.KKYFB&&window.KKYFB._stores?window.KKYFB._stores.families:null;
  if(!store||!store.rows)return 'missing-family-store';
  function findItem(name){
    var found=null,count=0;
    for(var i=0;i<store.rows.length;i++){
      var item=store.rows[i],attrs=item&&item.attrs?item.attrs:{};
      if(String(attrs['data-name']||'')==name){found=item;count++;}
    }
    return {item:found,count:count};
  }
  function text(value){return String(value||'');}
  function detailModalText(){
    var notes=document.getElementById('detailNotes');
    var buttons=notes&&notes.getElementsByClassName?notes.getElementsByClassName('fingerprint-diff-toggle'):[];
    if(!buttons||buttons.length<1)return 'NO_BUTTON';
    try{buttons[0].click();}catch(ex){return 'CLICK_ERROR '+(ex.message||ex.description||ex);}
    var mask=document.getElementById('diffModalMask'),body=document.getElementById('diffModalBody');
    var bodyText=body?text(body.innerText||body.textContent):'';
    if(!mask||String(mask.style.display||'').toLowerCase()!='block')bodyText='NOT_OPEN '+bodyText;
    if(mask)mask.style.display='none';
    return bodyText;
  }
  var parentResult=findItem('AUDIT_COMPOSITE_PARENT');
  var childResult=findItem('AUDIT_NESTED_FLOW_BOX');
  var missingResult=findItem('AUDIT_NESTED_MISSING_CHILD');
  var matchingResult=findItem('AUDIT_MATCHING_NESTED_CHILD');
  if(parentResult.count!=1)out.push('nested parent row count expected 1 but got '+parentResult.count);
  if(childResult.count!=1)out.push('differing nested child row count expected 1 but got '+childResult.count);
  if(missingResult.count!=1)out.push('missing nested child row count expected 1 but got '+missingResult.count);
  if(matchingResult.count!=0)out.push('matching nested child should stay hidden but row count is '+matchingResult.count);
  var parent=parentResult.item,child=childResult.item,missing=missingResult.item;
  if(child){
    var childAttrs=child.attrs||{},childDiff=text(childAttrs['data-diffrows']);
    if(text(childAttrs['data-nested-child']).toLowerCase()!='true')out.push('differing nested child lost data-nested-child=true');
    if(child.selectable)out.push('differing nested child remains directly selectable');
    if(childDiff.indexOf('Width=600')<0||childDiff.indexOf('Width=500')<0)out.push('differing nested child exact Width values missing: '+childDiff);
    var childAction=text(childAttrs['data-action']);
    if(childAction.indexOf('Review parent family')<0&&childAction.indexOf('상위 패밀리 검토')<0)out.push('differing nested child action is not parent review: '+childAction);
  }
  if(missing){
    var missingAttrs=missing.attrs||{},missingStatus=text(missingAttrs['data-status']),missingAction=text(missingAttrs['data-action']),missingNotes=text(missingAttrs['data-notes']),missingChanges=text(missingAttrs['data-changes']);
    if(text(missingAttrs['data-raw']).toLowerCase()!='nestedmissingfromparent')out.push('missing nested child raw status is wrong: '+text(missingAttrs['data-raw']));
    if(missingStatus.indexOf('Nested Family Missing')<0&&missingStatus.indexOf('하위 패밀리 누락')<0)out.push('missing nested child status label is wrong: '+missingStatus);
    if(missingAction.indexOf('Update parent family')<0&&missingAction.indexOf('상위 패밀리 업데이트 필요')<0)out.push('missing nested child action is wrong: '+missingAction);
    if(missingNotes.indexOf('AUDIT_MISSING_COMPOSITE_PARENT')<0)out.push('missing nested child memo lost its parent name: '+missingNotes);
    if((missingNotes+' '+missingChanges).indexOf('Fingerprint 생성 실패')>=0||(missingNotes+' '+missingChanges).indexOf('fingerprint was not created')>=0)out.push('missing nested child was mislabeled as a fingerprint capture failure');
    if(missing.selectable)out.push('missing nested child remains directly selectable');
  }
  if(parent){
    var parentAttrs=parent.attrs||{},parentDiff=text(parentAttrs['data-diffrows']);
    if(text(parentAttrs['data-raw']).toLowerCase()!='differentfromstandard')out.push('nested parent is not marked DifferentFromStandard: '+text(parentAttrs['data-raw']));
    if(parentDiff.indexOf('AUDIT_NESTED_FLOW_BOX')<0||parentDiff.indexOf('Width=600')<0||parentDiff.indexOf('Width=500')<0)out.push('nested parent exact child difference missing: '+parentDiff);
  }
  try{
    var search=document.getElementById('searchBox');if(search)search.value='';
    window.currentFilter='All';window.currentDiscipline='All';window.currentTreeDiscipline='All';window.currentTreeGroup='';window.currentTreeCategory='';
    window.advStatus='All';window.advGroup='All';window.advCategory='';window.advMismatchOnly=false;
    if(window.filterRows)window.filterRows('nested-family-audit');
  }catch(ignoreFilter){}
  if(child&&window.KKYFB&&window.KKYFB.findSavedRow){
    var childRow=window.KKYFB.findSavedRow('families',{name:'AUDIT_NESTED_FLOW_BOX',category:text(child.attrs['data-category']),kind:'',discipline:text(child.attrs['data-discipline-key'])});
    if(!childRow){
      out.push('differing nested child was not rendered');
    }else{
      var inputs=childRow.getElementsByTagName('input');
      if(!inputs.length||!inputs[0].disabled)out.push('differing nested child checkbox is not disabled');
      if(window.selectRow)window.selectRow(childRow,false);
      var childName=document.getElementById('detailName');
      if(!childName||text(childName.innerText||childName.textContent).indexOf('AUDIT_NESTED_FLOW_BOX')<0)out.push('nested child detail did not activate');
      var childModal=detailModalText();
      if(childModal.indexOf('Width=600')<0||childModal.indexOf('Width=500')<0)out.push('nested child detail table lost exact Width difference: '+childModal);
    }
  }
  if(parent&&window.KKYFB&&window.KKYFB.findSavedRow){
    var parentRow=window.KKYFB.findSavedRow('families',{name:'AUDIT_COMPOSITE_PARENT',category:text(parent.attrs['data-category']),kind:'',discipline:text(parent.attrs['data-discipline-key'])});
    if(!parentRow){
      out.push('nested parent was not rendered');
    }else{
      if(window.selectRow)window.selectRow(parentRow,false);
      var parentModal=detailModalText();
      if(parentModal.indexOf('AUDIT_NESTED_FLOW_BOX')<0||parentModal.indexOf('Width=600')<0||parentModal.indexOf('Width=500')<0)out.push('nested parent detail table lost propagated child difference: '+parentModal);
    }
  }
  return out.length?'FAIL '+out.join(' | '):'OK';
})()
");
		if (!string.Equals(state, "OK", StringComparison.OrdinalIgnoreCase))
		{
			result.Failures.Add("Nested family difference state check failed: " + state);
		}
	}

	private static List<string> CollectVisibleTextSnippets(WebBrowser browser)
	{
		string script = @"
(function(){
  var out=[];
  function compact(s){return String(s||'').replace(/\s+/g,' ').replace(/^\s+|\s+$/g,'');}
  function cls(el){return el?(' '+String(el.className||'')+' '):'';}
  function css(el,name){try{var s=(window.getComputedStyle?window.getComputedStyle(el,null):el.currentStyle);return s?(s[name]||s.getPropertyValue&&s.getPropertyValue(name)||''):'';}catch(e){return '';}}
  function visible(el){
    if(!el||el.nodeType!=1)return true;
    var tag=String(el.tagName||'').toLowerCase();
    if(tag=='script'||tag=='style'||tag=='noscript')return false;
    var c=cls(el);
    if(c.indexOf(' hidden ')>=0||c.indexOf(' sr-only ')>=0)return false;
    if(css(el,'display')=='none'||css(el,'visibility')=='hidden')return false;
    if(el.offsetWidth===0&&el.offsetHeight===0&&tag!='body'&&tag!='html')return false;
    return true;
  }
  function skipTextContainer(el){
    while(el&&el.nodeType==1){
      var id=String(el.id||'');
      var c=cls(el);
      if(id=='fbDebug'||id=='fbDebugLog')return true;
      if(c.indexOf(' path ')>=0||c.indexOf(' technical ')>=0||c.indexOf(' preview-diagnostic ')>=0)return true;
      el=el.parentNode;
    }
    return false;
  }
  function walk(n){
    if(!n||out.length>260)return;
    if(n.nodeType==3){
      var parent=n.parentNode;
      if(!visible(parent)||skipTextContainer(parent))return;
      var t=compact(n.nodeValue);
      if(t)out.push(t);
      return;
    }
    if(n.nodeType!=1||!visible(n))return;
    for(var child=n.firstChild;child;child=child.nextSibling)walk(child);
  }
  walk(document.body);
  return out.join('\u001f');
})()
";
		string value = EvalString(browser, script);
		if (string.IsNullOrWhiteSpace(value) || value.StartsWith("eval-error:", StringComparison.OrdinalIgnoreCase))
		{
			return new List<string>();
		}
		return value
			.Split(new[] { '\u001f' }, StringSplitOptions.RemoveEmptyEntries)
			.Select(CompactLanguageSnippet)
			.Where(snippet => !string.IsNullOrWhiteSpace(snippet))
			.ToList();
	}

	private static bool IsAllowedEnglishModeHangulSnippet(string snippet)
	{
		string text = CompactLanguageSnippet(snippet);
		if (string.Equals(text, "한국어", StringComparison.Ordinal))
		{
			return true;
		}
		if (text.IndexOf("한국어", StringComparison.Ordinal) >= 0 && HangulRegex.Matches(text).Count <= 3)
		{
			return true;
		}
		return false;
	}

	private static bool ContainsDisallowedEnglishUiInKoreanMode(string snippet)
	{
		string text = CompactLanguageSnippet(snippet);
		if (string.IsNullOrWhiteSpace(text) || !LatinRegex.IsMatch(text))
		{
			return false;
		}
		if (IsAllowedKoreanModeLatinSnippet(text))
		{
			return false;
		}
		foreach (string phrase in KoreanModeDisallowedEnglishPhrases)
		{
			if (text.IndexOf(phrase, StringComparison.OrdinalIgnoreCase) >= 0)
			{
				return true;
			}
		}
		return false;
	}

	private static bool IsAllowedKoreanModeLatinSnippet(string text)
	{
		string normalized = CompactLanguageSnippet(text);
		if (string.IsNullOrWhiteSpace(normalized))
		{
			return true;
		}
		string lower = normalized.ToLowerInvariant();
		if (lower.IndexOf(@":\", StringComparison.Ordinal) >= 0 ||
			lower.IndexOf("://", StringComparison.Ordinal) >= 0 ||
			lower.EndsWith(".rvt", StringComparison.Ordinal) ||
			lower.EndsWith(".rfa", StringComparison.Ordinal) ||
			lower.EndsWith(".xlsx", StringComparison.Ordinal) ||
			lower.EndsWith(".png", StringComparison.Ordinal))
		{
			return true;
		}
		string[] allowedFragments = new[]
		{
			"KKY",
			"RVT",
			"Revit",
			"Excel",
			"CSV",
			"JSON",
			"PDF",
			"PNG",
			"URL",
			"ID",
			"F12",
			"ON",
			"OFF",
			"UI_Audit",
			"UI audit",
			"UI_AUDIT",
			"AUDIT_",
			"Project1",
			"Family Browser"
		};
		foreach (string fragment in allowedFragments)
		{
			if (normalized.IndexOf(fragment, StringComparison.OrdinalIgnoreCase) >= 0)
			{
				return true;
			}
		}
		if (Regex.IsMatch(normalized, @"^[A-Z0-9_\- /().:]+$"))
		{
			return true;
		}
		return false;
	}

	private static string CompactLanguageSnippet(string value)
	{
		if (string.IsNullOrWhiteSpace(value))
		{
			return string.Empty;
		}
		string compact = Regex.Replace(value, @"\s+", " ").Trim();
		return compact.Length > 120 ? compact.Substring(0, 117) + "..." : compact;
	}

	private static void CheckAutoDetachedDetailAction(Dictionary<string, string> options, AuditResult result)
	{
		string activeTab = GetOption(options, "activeTab", string.Empty);
		bool browserTab = string.Equals(activeTab, "families", StringComparison.OrdinalIgnoreCase) ||
			string.Equals(activeTab, "systems", StringComparison.OrdinalIgnoreCase);
		if (!browserTab ||
			!GetBool(options, "includeRows", true) ||
			!GetBool(options, "standardRvtRegistered", true) ||
			!GetBool(options, "standardListRegistered", true))
		{
			return;
		}

		bool emitted = WaitUntil(delegate
		{
			return result.HostActions.Any(action => string.Equals(action, "detail-window-open", StringComparison.OrdinalIgnoreCase));
		}, 2500);

		if (!emitted)
		{
			result.Failures.Add("Auto detached detail action missing for browser tab '" + activeTab + "'. Expected host action detail-window-open when Family/System list has visible rows.");
		}
	}

	private static void CheckBrowserDetailContent(WebBrowser browser, Dictionary<string, string> options, AuditResult result)
	{
		if (!string.Equals(GetOption(options, "activeTab", string.Empty), "families", StringComparison.OrdinalIgnoreCase) ||
			!GetBool(options, "includeRows", true) ||
			!GetBool(options, "standardRvtRegistered", true) ||
			!GetBool(options, "standardListRegistered", true))
		{
			return;
		}

		string selectResult = EvalScript(browser, @"
(function(){
  if(!window.selectRow)return 'FAIL selectRow missing';
  var table=document.getElementById('familiesTable');
  if(!table)return 'SKIP familiesTable missing';
  var rows=table.getElementsByTagName('tr');
  var row=null;
  for(var i=0;i<rows.length;i++){
    var cls=rows[i].className||'';
    if(cls.indexOf('data')>=0 && rows[i].style.display!='none'){row=rows[i];break;}
  }
  if(!row)return 'SKIP no visible family data row';
  window.selectRow(row,false);
  return 'OK';
})()
");
		if (selectResult.StartsWith("SKIP", StringComparison.OrdinalIgnoreCase))
		{
			result.Warnings.Add("Detail content check skipped: " + selectResult);
			return;
		}
		if (!string.Equals(selectResult, "OK", StringComparison.OrdinalIgnoreCase))
		{
			result.Failures.Add("Detail content row selection failed: " + selectResult);
			return;
		}

		bool adminMode = GetBool(options, "adminMode", true);
		string adminLiteral = adminMode ? "true" : "false";
		string validationScript = @"
(function(){
  function txt(id){var e=document.getElementById(id);return e?(e.innerText||e.textContent||''):'';}
  var out=[];
  var name=txt('detailName');
  if(name.indexOf('AUDIT_SUPPLY_DIFFUSER')<0)out.push('detailName missing selected audit family: '+name);
  var category=txt('detailCategory');
  if(category.indexOf('Mechanical Equipment')<0)out.push('detailCategory missing selected category: '+category);
  var nested=txt('detailNested');
  if(nested.indexOf('AUDIT_SUPPLY_DIFFUSER_FACE')<0 || nested.indexOf('Air Terminals')<0)out.push('composite nested child table missing category/family: '+nested);
  if(nested.indexOf('AUDIT_FLOW_BOX')<0 || nested.indexOf('Mechanical Equipment')<0)out.push('composite nested child second row missing category/family: '+nested);
  var nestedTables=document.getElementsByClassName?document.getElementsByClassName('nested-child-table'):[];
  if(!nestedTables || nestedTables.length<1)out.push('nested child table markup missing');
  if(nestedTables && nestedTables.length>0){
    var nestedRows=nestedTables[0].getElementsByTagName('tr');
    var nestedBodyRowCount=0;
    var dashDuplicate=false;
    for(var nr=1;nr<nestedRows.length;nr++){
      var nestedCells=nestedRows[nr].getElementsByTagName('td');
      if(nestedCells.length<2)continue;
      nestedBodyRowCount++;
      var nestedCategory=(nestedCells[0].innerText||nestedCells[0].textContent||'').replace(/^\s+|\s+$/g,'');
      var nestedFamily=(nestedCells[1].innerText||nestedCells[1].textContent||'').replace(/^\s+|\s+$/g,'');
      if(nestedCategory=='-' && (nestedFamily=='AUDIT_SUPPLY_DIFFUSER_FACE' || nestedFamily=='AUDIT_FLOW_BOX'))dashDuplicate=true;
    }
    if(dashDuplicate)out.push('nested child duplicate with dash category was not removed');
    if(nestedBodyRowCount!=2)out.push('nested child table row count expected 2 after dedupe but got '+nestedBodyRowCount);
  }
  if(__ADMIN_MODE__){
    var params=txt('detailParameters');
    if(params.indexOf('Width')<0 || params.indexOf('600')<0)out.push('parameter detail missing Width/600: '+params);
    if(params.indexOf('KKY')<0)out.push('parameter detail missing common family value KKY: '+params);
    if(params.indexOf('SUP-01')<0)out.push('parameter detail missing instance value SUP-01: '+params);
    if(params.indexOf('Audit_SizeTable')<0 || params.indexOf('audit_sizetable')>=0 || (params.indexOf('12 rows x 5 columns')<0 && params.indexOf('12행 x 5열')<0))out.push('lookup CSV summary lost original table-name casing or size: '+params);
    var parameterTarget=document.getElementById('detailParameters');
    var unifiedTables=parameterTarget&&parameterTarget.getElementsByClassName?parameterTarget.getElementsByClassName('unified-parameter-table'):[];
    if(!unifiedTables || unifiedTables.length!=1)out.push('parameter detail expected exactly one unified table but got '+(unifiedTables?unifiedTables.length:0));
    var typeSelects=parameterTarget&&parameterTarget.getElementsByClassName?parameterTarget.getElementsByClassName('parameter-type-select'):[];
    if(!typeSelects || typeSelects.length!=1)out.push('parameter detail expected exactly one type selector but got '+(typeSelects?typeSelects.length:0));
    else{
      var typeSelect=typeSelects[0];
      if(typeSelect.options.length!=2)out.push('parameter type selector expected 2 audit types but got '+typeSelect.options.length);
      function readUnifiedParameterState(){
        var tables=parameterTarget.getElementsByClassName('unified-parameter-table');
		var state={tableCount:tables.length,widthCount:0,widthValue:'',airflowValue:'',yesNoValue:'',hasInstance:false};
        if(tables.length!=1)return state;
        var rows=tables[0].getElementsByTagName('tr');
        for(var pr=1;pr<rows.length;pr++){
          var cells=rows[pr].getElementsByTagName('td');if(cells.length<5)continue;
          var scope=(cells[1].innerText||cells[1].textContent||'').replace(/^\s+|\s+$/g,'');
          var name=(cells[2].innerText||cells[2].textContent||'').replace(/^\s+|\s+$/g,'');
          var value=(cells[3].innerText||cells[3].textContent||'').replace(/^\s+|\s+$/g,'');
          var isType=scope.toLowerCase()=='type'||scope.indexOf('타입')>=0;
          var isInstance=scope.toLowerCase()=='instance'||scope.indexOf('인스턴스')>=0;
		  if(isType&&name=='Width'){state.widthCount++;state.widthValue=value;}
		  if(isType&&name=='Airflow')state.airflowValue=value;
		  if(isType&&name=='IsEnabled')state.yesNoValue=value;
		  if(isInstance&&value=='SUP-01')state.hasInstance=true;
        }
        return state;
      }
		var initialState=readUnifiedParameterState();
		if(initialState.tableCount!=1 || initialState.widthCount!=1 || initialState.widthValue!='600')out.push('initial unified parameter table duplicated or lost selected-type Width: '+initialState.tableCount+'/'+initialState.widthCount+'/'+initialState.widthValue);
		if(initialState.yesNoValue!='Yes')out.push('Yes/No parameter expected Yes for the first type but got: '+initialState.yesNoValue);
		if(!initialState.hasInstance)out.push('initial unified parameter table lost instance row');
      if(typeSelect.options.length>1){
        typeSelect.selectedIndex=1;typeSelect.value='1';
        if(window.onParameterTypeChange)window.onParameterTypeChange(typeSelect);else out.push('onParameterTypeChange missing');
        var switchedState=readUnifiedParameterState();
		  if(switchedState.tableCount!=1 || switchedState.widthCount!=1 || switchedState.widthValue!='1200')out.push('type switch did not update the single unified Width row: '+switchedState.tableCount+'/'+switchedState.widthCount+'/'+switchedState.widthValue);
		  if(switchedState.airflowValue!='900 CFM')out.push('type switch did not update Airflow to 900 CFM: '+switchedState.airflowValue);
		  if(switchedState.yesNoValue!='No')out.push('Yes/No parameter expected No after type switch but got: '+switchedState.yesNoValue);
		  if(!switchedState.hasInstance)out.push('type switch removed the common instance row');
      }
    }
  }
  var selectedRaw='';
  var familyTable=document.getElementById('familiesTable');
  if(familyTable){
    var selectedRows=familyTable.getElementsByTagName('tr');
    for(var sr=0;sr<selectedRows.length;sr++){
      var selectedClass=selectedRows[sr].className||'';
      if(selectedClass.indexOf('data')>=0 && selectedClass.indexOf('selected')>=0){
        selectedRaw=selectedRows[sr].getAttribute('data-raw')||'';
        break;
      }
    }
  }
  if(selectedRaw!='LoadAvailable')out.push('selected audit family should be LoadAvailable for standard-detail validation: '+selectedRaw);
  var detailNotes=document.getElementById('detailNotes');
  var notesText=detailNotes?(detailNotes.innerText||detailNotes.textContent||''):'';
  var diffButtons=detailNotes&&detailNotes.getElementsByClassName?detailNotes.getElementsByClassName('fingerprint-diff-toggle'):[];
  if(diffButtons && diffButtons.length>0)out.push('load available detail should not show fingerprint diff detail button');
  if(notesText.indexOf('타입 수')>=0 || notesText.toLowerCase().indexOf('type count')>=0)out.push('load available detail should not show fingerprint diff summary: '+notesText);
  var preview=document.getElementById('preview');
  if(!preview){
    out.push('preview element missing');
  }else{
    var html=preview.innerHTML||'';
    var previewText=preview.innerText||preview.textContent||'';
    if(previewText.indexOf('불러오는')>=0 || previewText.toLowerCase().indexOf('loading')>=0)return 'WAIT preview loading';
    if(previewText.indexOf('No cached')>=0 || previewText.indexOf('캐시된 3D 이미지 없음')>=0)out.push('preview is showing no-cache fallback: '+previewText);
    if(html.indexOf('preview-fit-image')<0 && html.indexOf('data:image/png')<0 && html.indexOf('file:///')<0)out.push('preview image markup missing: '+previewText);
    var chips=preview.getElementsByClassName?preview.getElementsByClassName('preview-open-chip'):[];
    if(!chips || chips.length<1){
      out.push('preview large-view chip missing');
    }else{
      try{
        chips[0].click();
        var mask=document.getElementById('previewModalMask');
        if(!mask || String(mask.style.display||'').toLowerCase()!='block')out.push('preview large-view chip did not open modal');
        if(mask)mask.style.display='none';
      }catch(ex){
        out.push('preview large-view chip click failed: '+(ex.message||ex.description||ex));
      }
    }
  }
  if(familyTable && window.selectRow){
	if(window.KKYFB&&window.filterRows){
	  window.currentFilter='All';
	  window.currentDiscipline='All';
	  window.currentTreeDiscipline='All';
	  window.currentTreeGroup='';
	  window.currentTreeCategory='';
	  window.advStatus='All';
	  window.advGroup='All';
	  window.advCategory='';
	  window.advMismatchOnly=false;
	  window.filterRows('search');
	  familyTable=document.getElementById('familiesTable');
	}
    var diffRow=null;
    var diffRows=familyTable.getElementsByTagName('tr');
    for(var dr=0;dr<diffRows.length;dr++){
      var cls=diffRows[dr].className||'';
      var raw=diffRows[dr].getAttribute('data-raw')||'';
      var diffRaw=diffRows[dr].getAttribute('data-diffrows')||'';
      if(cls.indexOf('data')>=0 && raw!='LoadAvailable' && diffRaw.replace(/^\s+|\s+$/g,'').length>0){
        diffRow=diffRows[dr];
        break;
      }
    }
    if(!diffRow){
      out.push('fingerprint diff audit row missing');
    }else{
      window.selectRow(diffRow,false);
      var diffNotes=document.getElementById('detailNotes');
      var diffNotesText=diffNotes?(diffNotes.innerText||diffNotes.textContent||''):'';
      if(diffNotesText.indexOf('타입 수')<0 && diffNotesText.toLowerCase().indexOf('type count')<0){
        out.push('fingerprint diff concise summary missing type-count text: '+diffNotesText);
      }
      var modalButtons=diffNotes&&diffNotes.getElementsByClassName?diffNotes.getElementsByClassName('fingerprint-diff-toggle'):[];
      if(!modalButtons || modalButtons.length<1){
        out.push('fingerprint diff detail button missing');
      }else{
        try{
          modalButtons[0].click();
          var diffMask=document.getElementById('diffModalMask');
          var diffBody=document.getElementById('diffModalBody');
          var diffBodyText=diffBody?(diffBody.innerText||diffBody.textContent||''):'';
          var diffBodyHtml=diffBody?(diffBody.innerHTML||''):'';
          if(!diffMask || String(diffMask.style.display||'').toLowerCase()!='block')out.push('fingerprint diff detail button did not open modal');
          if(diffBodyHtml.indexOf('fingerprint-diff-table')<0)out.push('fingerprint diff modal table markup missing: '+diffBodyText);
          var hasTypeCount=diffBodyText.indexOf('Type Count')>=0 || diffBodyText.indexOf('타입 수')>=0;
          var hasLookupCsv=(diffBodyText.indexOf('Lookup CSV')>=0 || diffBodyText.indexOf('CSV 테이블')>=0) && diffBodyText.indexOf('Audit_SizeTable')>=0 && diffBodyText.indexOf('audit_sizetable')<0 && (diffBodyText.indexOf('12 rows x 5 columns')>=0 || diffBodyText.indexOf('12행 x 5열')>=0);
          if(!hasTypeCount || diffBodyText.indexOf('Width')<0 || !hasLookupCsv)out.push('fingerprint diff modal rows missing expected data: '+diffBodyText);
          if(diffMask)diffMask.style.display='none';
        }catch(ex){
          out.push('fingerprint diff detail button click failed: '+(ex.message||ex.description||ex));
        }
      }
    }
  }
  return out.length?'FAIL '+out.join(' | '):'OK';
})()
".Replace("__ADMIN_MODE__", adminLiteral);

		string validation = string.Empty;
		WaitUntil(delegate
		{
			DoEventsFor(80);
			validation = EvalScript(browser, validationScript);
			return string.Equals(validation, "OK", StringComparison.OrdinalIgnoreCase) ||
				validation.StartsWith("FAIL", StringComparison.OrdinalIgnoreCase);
		}, 3500);

		if (!string.Equals(validation, "OK", StringComparison.OrdinalIgnoreCase))
		{
			result.Failures.Add("Detail content check failed: " + validation);
		}
	}

	private static void CheckBrowserSystemDetailContent(WebBrowser browser, Dictionary<string, string> options, AuditResult result)
	{
		if (!string.Equals(GetOption(options, "activeTab", string.Empty), "systems", StringComparison.OrdinalIgnoreCase) ||
			!GetBool(options, "includeRows", true) ||
			!GetBool(options, "standardRvtRegistered", true) ||
			!GetBool(options, "standardListRegistered", true))
		{
			return;
		}

		bool expectDetailedSystemTypeComponents = GetBool(options, "compareDetailedSystemTypeComponents", true);
		string validation = EvalScript(browser, @"
(function(){
  function txt(id){var e=document.getElementById(id);return e?(e.innerText||e.textContent||''):'';}
  var out=[];
	var expectDetailedComponents=" + (expectDetailedSystemTypeComponents ? "true" : "false") + @";
  try{
    if(window.currentTab!='systems')return 'SKIP not on systems tab';
    var table=document.getElementById('systemsTable');
    if(!table)return 'SKIP systemsTable missing';
    var rows=table.getElementsByTagName('tr');
    var row=null;
    for(var i=0;i<rows.length;i++){
      var cls=rows[i].className||'';
      var name=rows[i].getAttribute('data-name')||'';
      if(cls.indexOf('data')>=0 && name.indexOf('AUDIT_SUPPLY_AIR')>=0){row=rows[i];break;}
    }
    if(!row)return 'FAIL system detail audit row missing';
    if(!window.selectRow)return 'FAIL selectRow missing for system detail';
    window.selectRow(row,false,true);
    var target=document.getElementById('detailParameters');
    var body=target?(target.innerText||target.textContent||''):'';
    var html=target?(target.innerHTML||''):'';
    var tables=target?target.getElementsByTagName('table'):[];
    var routingTables=[];
	var criteriaTables=[];
	var layerTables=[];
	var unitSelects=[];
	var boundedCriteriaCells=[];
	var unboundedCriteriaCells=[];
	var layerCoreBoundaries=[];
    var groupRows=[];
    var ruleRows=[];
    for(var t=0;t<tables.length;t++){
      var tc=tables[t].className||'';
      if(tc.indexOf('system-routing-preference-table')>=0)routingTables.push(tables[t]);
	  if(tc.indexOf('system-routing-criteria-table')>=0)criteriaTables.push(tables[t]);
	  if(tc.indexOf('system-layer-composition-table')>=0)layerTables.push(tables[t]);
    }
    if(target){
      var trs=target.getElementsByTagName('tr');
      for(var r=0;r<trs.length;r++){
        var rc=trs[r].className||'';
        if(rc.indexOf('system-routing-group-row')>=0)groupRows.push(trs[r]);
        if(rc.indexOf('system-routing-rule-row')>=0)ruleRows.push(trs[r]);
		if(rc.indexOf('system-layer-core-boundary')>=0)layerCoreBoundaries.push(trs[r]);
      }
	  var selects=target.getElementsByTagName('select');
	  for(var s=0;s<selects.length;s++){
		if((' '+(selects[s].className||'')+' ').indexOf(' system-routing-unit-select ')>=0)unitSelects.push(selects[s]);
	  }
    }
    if(routingTables.length!=1)out.push('system routing preference table expected exactly one, actual '+routingTables.length+': '+body);
	if(layerTables.length!=1)out.push('system layer composition table expected exactly one, actual '+layerTables.length+': '+body);
    if(groupRows.length<2)out.push('system routing preference groups missing: '+body);
    if(ruleRows.length!=2)out.push('system routing preference rule row count expected 2, actual '+ruleRows.length+': '+body);
	if(unitSelects.length!=2)out.push('system routing/layer unit selectors expected exactly two, actual '+unitSelects.length+': '+body);
	else if(unitSelects[0].value!='mm'||unitSelects[1].value!='mm')out.push('system routing/layer unit selectors default is not mm: '+unitSelects[0].value+' / '+unitSelects[1].value);
	if(layerCoreBoundaries.length<2)out.push('system layer core boundaries are incomplete: '+body);
	if(body.indexOf('AUDIT_BRICK')<0||body.indexOf('AUDIT_CONCRETE')<0||body.indexOf('AUDIT_GYPSUM')<0)out.push('system layer material rows are incomplete: '+body);
	if(body.indexOf('Exterior')<0&&body.indexOf('외부')<0)out.push('system layer exterior direction is missing: '+body);
	if(body.indexOf('Interior')<0&&body.indexOf('내부')<0)out.push('system layer interior direction is missing: '+body);
	if((body.indexOf('Structural material')<0&&body.indexOf('구조 재료')<0)||(body.indexOf('Variable')<0&&body.indexOf('가변')<0))out.push('system layer structural/variable badges are missing: '+body);
	if(body.indexOf('200 mm')<0||body.indexOf('15 mm')<0)out.push('system layer mm conversion missing: '+body);
	if(criteriaTables.length!=1)out.push('system routing criteria table expected exactly one, actual '+criteriaTables.length+': '+body);
	else{
	  var criterionRows=criteriaTables[0].getElementsByTagName('tr');
	  if(criterionRows.length!=3)out.push('system routing criteria row count expected 2, actual '+(criterionRows.length-1)+': '+body);
	  else{
		boundedCriteriaCells=criterionRows[1].getElementsByTagName('td');
		unboundedCriteriaCells=criterionRows[2].getElementsByTagName('td');
		if(boundedCriteriaCells.length<3)out.push('bounded system routing criterion columns missing: '+body);
		else{
		  var minimumText=boundedCriteriaCells[1].innerText||boundedCriteriaCells[1].textContent||'';
		  var maximumText=boundedCriteriaCells[2].innerText||boundedCriteriaCells[2].textContent||'';
		  if(minimumText.indexOf('100 mm')<0)out.push('system routing minimum cell missing 100 mm: '+minimumText);
		  if(maximumText.indexOf('300 mm')<0)out.push('system routing maximum cell missing 300 mm: '+maximumText);
		}
		if(unboundedCriteriaCells.length<3)out.push('unbounded system routing criterion columns missing: '+body);
		else{
		  var unboundedMinimum=unboundedCriteriaCells[1].innerText||unboundedCriteriaCells[1].textContent||'';
		  var unboundedMaximum=unboundedCriteriaCells[2].innerText||unboundedCriteriaCells[2].textContent||'';
		  if((unboundedMinimum.indexOf('No limit')<0&&unboundedMinimum.indexOf('제한 없음')<0)||(unboundedMaximum.indexOf('No limit')<0&&unboundedMaximum.indexOf('제한 없음')<0))out.push('system routing unbounded minimum/maximum cells are incomplete: '+unboundedMinimum+' / '+unboundedMaximum);
		}
	  }
	}
    if(body.indexOf('AUDIT_DUCT_SEGMENT')<0)out.push('system detail segment missing: '+body);
    if(body.indexOf('12')<0)out.push('system routing size count missing: '+body);
	if(body.indexOf('100 mm')<0 || body.indexOf('300 mm')<0)out.push('system routing mm conversion missing: '+body);
	if(body.indexOf('No limit')<0 && body.indexOf('제한 없음')<0)out.push('system routing sentinel not rendered as no limit: '+body);
	if(body.toLowerCase().indexOf('e+30')>=0 || body.toLowerCase().indexOf('min=')>=0 || body.toLowerCase().indexOf('max=')>=0)out.push('system routing scientific sentinel leaked into visible text: '+body);
	if(unitSelects.length==2 && window.changeSystemRoutingUnit){
	  window.__auditPersistedMeasurementUnit='';
	  window.persistSystemRoutingUnit=function(unit){window.__auditPersistedMeasurementUnit=unit;};
	  unitSelects[0].value='in';
	  window.changeSystemRoutingUnit(unitSelects[0]);
	  var inchBody=target?(target.innerText||target.textContent||''):'';
	  if(inchBody.indexOf('3.937 in')<0 || inchBody.indexOf('11.811 in')<0)out.push('system routing inch conversion missing: '+inchBody);
	  if(inchBody.indexOf('7.874 in')<0||inchBody.indexOf('0.591 in')<0)out.push('system layer inch conversion missing: '+inchBody);
	  if(unitSelects[0].value!='in'||unitSelects[1].value!='in')out.push('system routing/layer inch selectors are not synchronized: '+unitSelects[0].value+' / '+unitSelects[1].value);
	  if(window.__auditPersistedMeasurementUnit!='in')out.push('system measurement unit change did not request inch persistence');
	  if(boundedCriteriaCells.length>=3){
		var minimumInches=boundedCriteriaCells[1].innerText||boundedCriteriaCells[1].textContent||'';
		var maximumInches=boundedCriteriaCells[2].innerText||boundedCriteriaCells[2].textContent||'';
		if(minimumInches.indexOf('3.937 in')<0)out.push('system routing minimum cell inch conversion missing: '+minimumInches);
		if(maximumInches.indexOf('11.811 in')<0)out.push('system routing maximum cell inch conversion missing: '+maximumInches);
	  }
	  unitSelects[1].value='mm';
	  window.changeSystemRoutingUnit(unitSelects[1]);
	  var restoredBody=target?(target.innerText||target.textContent||''):'';
	  if(unitSelects[0].value!='mm'||unitSelects[1].value!='mm')out.push('system routing/layer mm selectors are not synchronized after restore: '+unitSelects[0].value+' / '+unitSelects[1].value);
	  if(restoredBody.indexOf('200 mm')<0||restoredBody.indexOf('15 mm')<0)out.push('system layer mm restore failed: '+restoredBody);
	  if(window.__auditPersistedMeasurementUnit!='mm')out.push('system measurement unit change did not request mm persistence');
	}else out.push('system routing unit change handler missing');
    if(body.indexOf('AUDIT_DUCT_ELBOW')<0)out.push('system detail dependent family missing: '+body);
	var dashboardHtml=document.body?(document.body.innerHTML||''):'';
	var koreanMode=dashboardHtml.indexOf('kkyfb:lang-en')>=0;
	var expectedElbowLabel=koreanMode?'엘보 (Elbows)':'Elbows';
	if(body.indexOf(expectedElbowLabel)<0 || (!koreanMode && body.indexOf('엘보')>=0))out.push('system routing bilingual elbow label missing for '+(koreanMode?'ko':'en')+': '+body);
    var title=txt('detailParameterTitle');
    if(title.indexOf('Routing')<0 && title.indexOf('라우팅')<0)out.push('system routing preference title missing: '+title);
    if(html.indexOf('system-detail-section')>=0 ||
       body.indexOf('Segments / Sizes')>=0 || body.indexOf('세그먼트 / 사이즈')>=0 ||
       body.indexOf('Dependent Loadable Families')>=0 || body.indexOf('의존 로더블 패밀리')>=0 ||
       body.indexOf('System Type Review Data')>=0 || body.indexOf('시스템 타입 검토 데이터')>=0){
      out.push('system detail legacy split sections remain visible: '+body);
    }
    var basicTables=[];
    for(var b=0;b<tables.length;b++)if((tables[b].className||'').indexOf('system-routing-basic-table')>=0)basicTables.push(tables[b]);
    for(var bt=0;bt<basicTables.length;bt++){
      var basicRows=basicTables[bt].getElementsByTagName('tr');
      for(var br=0;br<basicRows.length;br++){
        var heads=basicRows[br].getElementsByTagName('th');
        var label=heads.length?(heads[0].innerText||heads[0].textContent||''):'';
        var key=label.replace(/\s+/g,'').toLowerCase();
        if(key=='segment'||key=='material'||key=='세그먼트'||key=='재료')out.push('system detail Segment/Material dash row remains visible: '+label);
      }
    }
    var previewBlock=document.getElementById('previewBlock');
    if(previewBlock && previewBlock.style.display!='none')out.push('system detail bottom review block is still visible');
	var railingRow=null;
	for(var rr=0;rr<rows.length;rr++){
	  var railingName=rows[rr].getAttribute('data-name')||'';
	  if(railingName.indexOf('AUDIT_GUARDRAIL')>=0){railingRow=rows[rr];break;}
	}
	if(!railingRow)out.push('detailed Railing component audit row missing');
	else{
	  window.selectRow(railingRow,false,true);
	  var componentBody=target?(target.innerText||target.textContent||''):'';
	  var componentTables=target?target.getElementsByTagName('table'):[];
	  var detailedTables=[];
	  for(var ct=0;ct<componentTables.length;ct++){
		if(componentTables[ct].getAttribute('data-system-component-table')=='true')detailedTables.push(componentTables[ct]);
	  }
	  if(expectDetailedComponents){
		if(detailedTables.length!=2)out.push('Railing detailed component/configuration difference tables expected 2, actual '+detailedTables.length+': '+componentBody);
		if(componentBody.indexOf('AUDIT_TOP_RAIL')<0||componentBody.indexOf('AUDIT_HANDRAIL')<0||componentBody.indexOf('AUDIT_BALUSTER')<0)out.push('Railing detailed component references are incomplete: '+componentBody);
		if(componentBody.indexOf('AUDIT_BALUSTER / Light')<0)out.push('Railing detailed component difference is missing: '+componentBody);
		if((componentBody.indexOf('Detailed Components')<0&&componentBody.indexOf('상세 구성 요소')<0)||(componentBody.indexOf('Component Differences')<0&&componentBody.indexOf('상세 구성 차이')<0))out.push('Railing detailed component table headings are missing: '+componentBody);
		var componentSelects=[];
		var componentAllSelects=target?target.getElementsByTagName('select'):[];
		for(var cs=0;cs<componentAllSelects.length;cs++)if((' '+(componentAllSelects[cs].className||'')+' ').indexOf(' system-component-unit-select ')>=0)componentSelects.push(componentAllSelects[cs]);
		if(componentSelects.length!=2)out.push('Railing component unit selectors expected 2, actual '+componentSelects.length+': '+componentBody);
		if(componentBody.indexOf('914.4 mm')<0||componentBody.indexOf('76.2 mm')<0||componentBody.indexOf('152.4 mm')<0)out.push('Railing component mm values are incomplete: '+componentBody);
		if(componentSelects.length==2&&window.changeSystemRoutingUnit){
		  componentSelects[0].value='in';window.changeSystemRoutingUnit(componentSelects[0]);
		  var componentInchBody=target?(target.innerText||target.textContent||''):'';
		  if(componentInchBody.indexOf('36 in')<0||componentInchBody.indexOf('3 in')<0||componentInchBody.indexOf('6 in')<0)out.push('Railing component inch conversion is incomplete: '+componentInchBody);
		  if(componentSelects[0].value!='in'||componentSelects[1].value!='in')out.push('Railing component unit selectors are not synchronized');
		  if(window.__auditPersistedMeasurementUnit!='in')out.push('Railing component inch selection did not request persistence');
		  componentSelects[1].value='mm';window.changeSystemRoutingUnit(componentSelects[1]);
		  if(componentSelects[0].value!='mm'||componentSelects[1].value!='mm')out.push('Railing component unit selectors did not restore mm');
		}
	  }else{
		if(detailedTables.length!=0)out.push('Railing detailed component tables remain visible while comparison is disabled: '+detailedTables.length);
		if(componentBody.indexOf('AUDIT_TOP_RAIL')>=0||componentBody.indexOf('AUDIT_HANDRAIL')>=0||componentBody.indexOf('AUDIT_BALUSTER')>=0)out.push('Railing detailed component references remain visible while comparison is disabled: '+componentBody);
	  }
	}
	var curtainRow=null;
	for(var cw=0;cw<rows.length;cw++){
	  var curtainName=rows[cw].getAttribute('data-name')||'';
	  if(curtainName.indexOf('AUDIT_CURTAIN_WALL')>=0){curtainRow=rows[cw];break;}
	}
	if(!curtainRow)out.push('mandatory curtain panel dependency audit row missing');
	else{
	  window.selectRow(curtainRow,false,true);
	  var curtainBody=target?(target.innerText||target.textContent||''):'';
	  var curtainTables=target?target.getElementsByTagName('table'):[];
	  var mandatoryCurtainTables=[];
	  for(var cwt=0;cwt<curtainTables.length;cwt++){
		if(curtainTables[cwt].getAttribute('data-system-curtain-panel-table')=='true')mandatoryCurtainTables.push(curtainTables[cwt]);
	  }
	  if(mandatoryCurtainTables.length!=2)out.push('curtain panel dependency/configuration difference tables expected 2, actual '+mandatoryCurtainTables.length+': '+curtainBody);
	  if(curtainBody.indexOf('AUDIT_SYSTEM_PANEL / Glazed')<0||curtainBody.indexOf('AUDIT_PANEL_SUPPORT / Standard')<0)out.push('curtain panel dependency references are incomplete: '+curtainBody);
	  if(curtainBody.indexOf('AUDIT_SYSTEM_PANEL / Solid')<0)out.push('curtain panel dependency difference is missing: '+curtainBody);
	  if((curtainBody.indexOf('Curtain Panel Dependencies')<0&&curtainBody.indexOf('커튼패널 의존 구성')<0)||(curtainBody.indexOf('Curtain Panel Differences')<0&&curtainBody.indexOf('커튼패널 구성 차이')<0))out.push('curtain panel dependency headings are missing: '+curtainBody);
	  var curtainUnitSelects=[];
	  var curtainAllSelects=target?target.getElementsByTagName('select'):[];
	  for(var cus=0;cus<curtainAllSelects.length;cus++)if((' '+(curtainAllSelects[cus].className||'')+' ').indexOf(' system-component-unit-select ')>=0)curtainUnitSelects.push(curtainAllSelects[cus]);
	  if(curtainUnitSelects.length!=2)out.push('curtain panel unit selectors expected 2, actual '+curtainUnitSelects.length+': '+curtainBody);
	  if(curtainBody.indexOf('1219.2 mm')<0||curtainBody.indexOf('1524 mm')<0)out.push('curtain panel mm values are incomplete: '+curtainBody);
	  if(curtainUnitSelects.length==2&&window.changeSystemRoutingUnit){
		curtainUnitSelects[0].value='in';window.changeSystemRoutingUnit(curtainUnitSelects[0]);
		var curtainInchBody=target?(target.innerText||target.textContent||''):'';
		if(curtainInchBody.indexOf('48 in')<0||curtainInchBody.indexOf('60 in')<0)out.push('curtain panel inch conversion is incomplete: '+curtainInchBody);
		if(curtainUnitSelects[0].value!='in'||curtainUnitSelects[1].value!='in')out.push('curtain panel unit selectors are not synchronized');
		if(window.__auditPersistedMeasurementUnit!='in')out.push('curtain panel inch selection did not request persistence');
		curtainUnitSelects[1].value='mm';window.changeSystemRoutingUnit(curtainUnitSelects[1]);
	  }
	}
	var panelTypeRow=null;
	for(var cp=0;cp<rows.length;cp++){
	  var panelTypeName=rows[cp].getAttribute('data-name')||'';
	  if(panelTypeName.indexOf('AUDIT_SYSTEM_PANEL_TYPE')>=0){panelTypeRow=rows[cp];break;}
	}
	if(!panelTypeRow)out.push('direct PanelType mandatory comparison audit row missing');
	else{
	  window.selectRow(panelTypeRow,false,true);
	  var panelTypeBody=target?(target.innerText||target.textContent||''):'';
	  var panelTypeTables=target?target.getElementsByTagName('table'):[];
	  var mandatoryPanelTypeTables=[];
	  for(var cpt=0;cpt<panelTypeTables.length;cpt++){
		if(panelTypeTables[cpt].getAttribute('data-system-curtain-panel-table')=='true')mandatoryPanelTypeTables.push(panelTypeTables[cpt]);
	  }
	  if(mandatoryPanelTypeTables.length!=2)out.push('direct PanelType dependency/difference tables expected 2, actual '+mandatoryPanelTypeTables.length+': '+panelTypeBody);
	  if(panelTypeBody.indexOf('AUDIT_SYSTEM_PANEL_TYPE / Glazed')<0||panelTypeBody.indexOf('AUDIT_PANEL_INSERT / Standard')<0)out.push('direct PanelType dependency references are incomplete: '+panelTypeBody);
	  if(panelTypeBody.indexOf('AUDIT_PANEL_INSERT / Light')<0)out.push('direct PanelType dependency difference is missing: '+panelTypeBody);
	  var panelUnitSelects=[];
	  var panelAllSelects=target?target.getElementsByTagName('select'):[];
	  for(var pus=0;pus<panelAllSelects.length;pus++)if((' '+(panelAllSelects[pus].className||'')+' ').indexOf(' system-component-unit-select ')>=0)panelUnitSelects.push(panelAllSelects[pus]);
	  if(panelUnitSelects.length!=2)out.push('direct PanelType unit selectors expected 2, actual '+panelUnitSelects.length+': '+panelTypeBody);
	  if(panelTypeBody.indexOf('76.2 mm')<0||panelTypeBody.indexOf('152.4 mm')<0)out.push('direct PanelType mm values are incomplete: '+panelTypeBody);
	  if(panelUnitSelects.length==2&&window.changeSystemRoutingUnit){
		panelUnitSelects[0].value='in';window.changeSystemRoutingUnit(panelUnitSelects[0]);
		var panelInchBody=target?(target.innerText||target.textContent||''):'';
		if(panelInchBody.indexOf('3 in')<0||panelInchBody.indexOf('6 in')<0)out.push('direct PanelType inch conversion is incomplete: '+panelInchBody);
		if(panelUnitSelects[0].value!='in'||panelUnitSelects[1].value!='in')out.push('direct PanelType unit selectors are not synchronized');
		if(window.__auditPersistedMeasurementUnit!='in')out.push('direct PanelType inch selection did not request persistence');
		panelUnitSelects[1].value='mm';window.changeSystemRoutingUnit(panelUnitSelects[1]);
	  }
	}
  }catch(ex){
    out.push('system detail check exception: '+(ex.message||ex.description||ex));
  }
  return out.length?'FAIL '+out.join(' | '):'OK';
})()
");
		if (validation.StartsWith("SKIP", StringComparison.OrdinalIgnoreCase))
		{
			result.Warnings.Add("System detail content check skipped: " + validation);
			return;
		}
		if (!string.Equals(validation, "OK", StringComparison.OrdinalIgnoreCase))
		{
			result.Failures.Add("System detail content check failed: " + validation);
		}
	}

	private static void CheckDetailedFilterResetAcrossBrowserTabs(WebBrowser browser, Dictionary<string, string> options, AuditResult result)
	{
		string activeTab = GetOption(options, "activeTab", string.Empty);
		if ((!string.Equals(activeTab, "families", StringComparison.OrdinalIgnoreCase) &&
			!string.Equals(activeTab, "systems", StringComparison.OrdinalIgnoreCase)) ||
			!GetBool(options, "includeRows", true) ||
			!GetBool(options, "standardRvtRegistered", true) ||
			!GetBool(options, "standardListRegistered", true))
		{
			return;
		}

		string outcome = EvalString(browser, @"
(function(){
  var failures=[];
  function seedDetailedFilter(){
    var status=document.getElementById('advStatus');
    var group=document.getElementById('advGroup');
    var category=document.getElementById('advCategory');
    var mismatch=document.getElementById('advMismatchOnly');
    if(!status||!group||!category)return 'detailed filter controls missing';
    status.value='LoadAvailable';
    group.value='Model';
    category.value='AUDIT_CROSS_TAB_FILTER';
    if(mismatch)mismatch.checked=true;
    if(typeof window.applyAdvancedFilter!=='function')return 'applyAdvancedFilter missing';
    window.applyAdvancedFilter();
    if(window.advStatus==='All'&&window.advGroup==='All'&&window.advCategory===''&&!window.advMismatchOnly)return 'detailed filter did not become non-default';
    return '';
  }
  function checkReset(expectedTab,label){
    var status=document.getElementById('advStatus');
    var group=document.getElementById('advGroup');
    var category=document.getElementById('advCategory');
    var mismatch=document.getElementById('advMismatchOnly');
    if(window.currentTab!==expectedTab)failures.push(label+' current tab='+window.currentTab);
    if(window.advStatus!=='All')failures.push(label+' advStatus='+window.advStatus);
    if(window.advGroup!=='All')failures.push(label+' advGroup='+window.advGroup);
    if(window.advCategory!=='')failures.push(label+' advCategory='+window.advCategory);
    if(!!window.advMismatchOnly)failures.push(label+' advMismatchOnly=true');
    if(status&&status.value!=='All')failures.push(label+' status control='+status.value);
    if(group&&group.value!=='All')failures.push(label+' group control='+group.value);
    if(category&&category.value!=='')failures.push(label+' category control='+category.value);
    if(mismatch&&mismatch.checked)failures.push(label+' mismatch control=true');
  }
  if(typeof window.setTab!=='function')return 'FAIL setTab missing';
  var start=String(window.currentTab||'');
  if(start!=='families'&&start!=='systems')return 'FAIL unexpected start tab '+start;
  var other=start==='families'?'systems':'families';
  if(!document.getElementById(start+'Pane'))return 'FAIL active browser pane missing';
  var otherPane=document.getElementById(other+'Pane');
  var createdOtherPane=false;
  if(!otherPane){
    var center=document.getElementById('mainCenter');
    if(!center)return 'FAIL mainCenter missing';
    otherPane=document.createElement('div');
    otherPane.id=other+'Pane';
    otherPane.className='pane';
    otherPane.style.display='none';
    center.appendChild(otherPane);
    createdOtherPane=true;
  }
  function cleanup(){if(createdOtherPane&&otherPane&&otherPane.parentNode)otherPane.parentNode.removeChild(otherPane);}
  var seedError=seedDetailedFilter();
  if(seedError){cleanup();return 'FAIL '+seedError;}
  window.setTab(other);
  checkReset(other,start+'->'+other);
  seedError=seedDetailedFilter();
  if(seedError){window.setTab(start);cleanup();return 'FAIL '+seedError+' after '+other;}
  window.setTab(start);
  checkReset(start,other+'->'+start);
  cleanup();
  return failures.length?'FAIL '+failures.join(' | '):'OK '+start+'<->'+other;
})()");
		if (!outcome.StartsWith("OK ", StringComparison.Ordinal))
		{
			result.Failures.Add("Detailed filter was not reset across Family/System tab switch: " + outcome);
		}
	}

	private static void CheckSearchFilteringKeepsFocus(WebBrowser browser, Dictionary<string, string> options, AuditResult result)
	{
		string activeTab = GetOption(options, "activeTab", string.Empty);
		bool browserTab = string.Equals(activeTab, "families", StringComparison.OrdinalIgnoreCase) ||
			string.Equals(activeTab, "systems", StringComparison.OrdinalIgnoreCase);
		if (!browserTab ||
			!GetBool(options, "includeRows", true) ||
			!GetBool(options, "standardRvtRegistered", true) ||
			!GetBool(options, "standardListRegistered", true))
		{
			return;
		}

		DoEventsFor(260);
		int hostActionBaseline = result.HostActions.Count;
		string setup = EvalScript(browser, @"
(function(){
  if(!window.queueFilterRows)return 'FAIL queueFilterRows missing';
  if(!window.filterRows)return 'FAIL filterRows missing';
  var search=document.getElementById('searchBox');
  if(!search)return 'SKIP searchBox missing';
  function pickSearchTerm(hay){
    hay=String(hay).replace(/[^\w\u3131-\u318e\uac00-\ud7a3]+/g,' ').replace(/^\s+|\s+$/g,'');
    var parts=hay.split(/\s+/);
    for(var i=0;i<parts.length;i++){
      if(parts[i].length>=2)return parts[i].length>12?parts[i].substring(0,12):parts[i];
    }
    return '';
  }
	  var table=document.getElementById((window.currentTab=='systems')?'systemsTable':'familiesTable');
	  if(!table)return 'SKIP browser table missing';
	  if(window.filterRowsTimer){clearTimeout(window.filterRowsTimer);window.filterRowsTimer=null;}
	  currentFilter='All';
	  currentDiscipline='All';
	  currentTreeDiscipline='All';currentTreeGroup='';currentTreeCategory='';
	  currentSystemTreeDiscipline='All';currentSystemTreeCategory='';
	  advStatus='All';advGroup='All';advCategory='';advMismatchOnly=false;
	  var advStatusControl=document.getElementById('advStatus');if(advStatusControl)advStatusControl.value='All';
	  var advGroupControl=document.getElementById('advGroup');if(advGroupControl)advGroupControl.value='All';
	  var advCategoryControl=document.getElementById('advCategory');if(advCategoryControl)advCategoryControl.value='';
	  var advMismatchControl=document.getElementById('advMismatchOnly');if(advMismatchControl)advMismatchControl.checked=false;
	  if(typeof window.setFilterChrome==='function')window.setFilterChrome('All');
	  if(typeof window.setDisciplineChrome==='function')window.setDisciplineChrome('All');
	  search.value='';
	  var store=window.KKYFB&&window.KKYFB._stores?window.KKYFB._stores[window.currentTab]:null;
	  if(store){store.page=0;store.filterSignature='';store.renderSignature='';}
	  window.filterRows('search');
	  var term='';
	  if(store&&store.rows&&store.rows.length){
	    term=String((store.rows[0].attrs&&store.rows[0].attrs['data-name'])||'').replace(/^\s+|\s+$/g,'');
	    if(!term)term=pickSearchTerm(store.rows[0].searchText||'');
	  }
  if(!term){
    var rows=table.getElementsByTagName('tr');
    for(var i=0;i<rows.length;i++){
      var cls=rows[i].className||'';
      if(cls.indexOf('data')<0 || rows[i].style.display=='none')continue;
      term=pickSearchTerm([
        rows[i].getAttribute('data-name')||'',
        rows[i].getAttribute('data-category')||'',
        rows[i].getAttribute('data-discipline')||'',
        rows[i].getAttribute('data-kind')||'',
        rows[i].getAttribute('data-status')||''
      ].join(' '));
      if(term)break;
    }
  }
	  if(!term)return 'SKIP no visible searchable data row';
	  search.readOnly=true;
	  search.value=term;
	  search.focus();
	  search.value=term;
	  if(window.filterRowsTimer){clearTimeout(window.filterRowsTimer);window.filterRowsTimer=null;}
	  if(store){store.page=0;store.filterSignature='';store.renderSignature='';}
	  window.queueFilterRows();
	  return 'OK '+term;
})()
");
		if (setup.StartsWith("SKIP", StringComparison.OrdinalIgnoreCase))
		{
			result.Warnings.Add("Search focus check skipped: " + setup);
			return;
		}
		if (!setup.StartsWith("OK", StringComparison.OrdinalIgnoreCase))
		{
			result.Failures.Add("Search focus setup failed: " + setup);
			return;
		}

		DoEventsFor(650);
		string validationScript = @"
(function(){
  function txt(id){var e=document.getElementById(id);return e?(e.innerText||e.textContent||''):'';}
  var out=[];
  var search=document.getElementById('searchBox');
  if(!search)return 'FAIL searchBox missing after queued search';
  if(document.activeElement!==search)out.push('search focus not retained');
	  var table=document.getElementById((window.currentTab=='systems')?'systemsTable':'familiesTable');
	  if(!table)out.push('browser table missing after queued search');
	  var store=window.KKYFB&&window.KKYFB._stores?window.KKYFB._stores[window.currentTab]:null;
	  var visible=0;
  var firstName='';
  if(table){
    var rows=table.getElementsByTagName('tr');
    for(var i=0;i<rows.length;i++){
      var cls=rows[i].className||'';
      if(cls.indexOf('data')<0 || rows[i].style.display=='none')continue;
      visible++;
      if(!firstName)firstName=rows[i].getAttribute('data-name')||'';
    }
  }
	  if(visible<1)out.push('queued search hid every data row (filtered='+(store&&store.filtered?store.filtered.length:'missing')+', query='+(search?search.value:'')+')');
  var detail=txt('detailName');
  if(firstName && detail.indexOf(firstName)<0)out.push('detail did not follow filtered first row: '+detail+' / '+firstName);
  return out.length?'FAIL '+out.join(' | '):'OK';
})()
";
		string validation = string.Empty;
		WaitUntil(delegate
		{
			DoEventsFor(80);
			validation = EvalScript(browser, validationScript);
			return string.Equals(validation, "OK", StringComparison.OrdinalIgnoreCase);
		}, 2500);
		if (!string.Equals(validation, "OK", StringComparison.OrdinalIgnoreCase))
		{
			result.Failures.Add("Search focus check failed: " + validation);
		}

		List<string> newHostActions = result.HostActions
			.Skip(hostActionBaseline)
			.Where(action => string.Equals(action, "detail-window-open", StringComparison.OrdinalIgnoreCase) ||
				string.Equals(action, "detail-window-sync", StringComparison.OrdinalIgnoreCase) ||
				action.StartsWith("preview-inline/", StringComparison.OrdinalIgnoreCase))
			.ToList();
		if (newHostActions.Count > 0)
		{
			result.Failures.Add("Search filtering emitted focus-stealing host actions: " + string.Join(", ", newHostActions.ToArray()));
		}
	}

	private static string EvalScript(WebBrowser browser, string script)
	{
		try
		{
			if (browser == null || browser.Document == null)
			{
				return "FAIL browser document missing";
			}
			return Convert.ToString(browser.Document.InvokeScript("eval", new object[] { script })) ?? string.Empty;
		}
		catch (Exception ex)
		{
			return "FAIL eval threw " + ex.GetType().Name + ": " + ex.Message;
		}
	}

	private static List<Clickable> CollectClickables(WebBrowser browser)
	{
		List<Clickable> items = new List<Clickable>();
		AddElements(items, browser.Document.GetElementsByTagName("a"), "a");
		AddElements(items, browser.Document.GetElementsByTagName("button"), "button");
		return items.Where(x => IsVisibleEnabled(x.Element)).ToList();
	}

	private static void AddElements(List<Clickable> items, HtmlElementCollection elements, string tag)
	{
		int index = 0;
		foreach (HtmlElement element in elements)
		{
			items.Add(new Clickable
			{
				Tag = tag,
				Text = Compact(element.InnerText),
				Href = NormalizeHref(element.GetAttribute("href")),
				OnClick = element.GetAttribute("onclick") ?? string.Empty,
				ClassName = element.GetAttribute("className") ?? string.Empty,
				Index = index++,
				Element = element
			});
		}
	}

	private static bool IsVisibleEnabled(HtmlElement element)
	{
		if (element == null)
		{
			return false;
		}
		string cls = (element.GetAttribute("className") ?? string.Empty).ToLowerInvariant();
		if (cls.IndexOf("disabled", StringComparison.Ordinal) >= 0)
		{
			return false;
		}
		HtmlElement current = element;
		while (current != null)
		{
			string style = (current.Style ?? string.Empty).ToLowerInvariant();
			if (style.IndexOf("display:none", StringComparison.Ordinal) >= 0 || style.IndexOf("visibility:hidden", StringComparison.Ordinal) >= 0)
			{
				return false;
			}
			current = current.Parent;
		}
		return element.OffsetRectangle.Width > 0 && element.OffsetRectangle.Height > 0;
	}

	private static string NormalizeHref(string href)
	{
		if (string.IsNullOrWhiteSpace(href))
		{
			return string.Empty;
		}
		string value = href.Trim();
		if (value.StartsWith("about:blank", StringComparison.OrdinalIgnoreCase))
		{
			value = value.Substring("about:blank".Length);
		}
		return value;
	}

	private static bool IsHostActionHref(string href)
	{
		if (string.IsNullOrWhiteSpace(href))
		{
			return false;
		}
		return href.StartsWith("kkyfb:", StringComparison.OrdinalIgnoreCase) || href.StartsWith("about:kkyfb:", StringComparison.OrdinalIgnoreCase);
	}

	private static string NormalizeHostAction(string raw)
	{
		if (string.IsNullOrWhiteSpace(raw))
		{
			return string.Empty;
		}
		if (raw.StartsWith("kkyfb:", StringComparison.OrdinalIgnoreCase))
		{
			return raw.Substring("kkyfb:".Length).Trim('/', ' ');
		}
		if (raw.StartsWith("about:kkyfb:", StringComparison.OrdinalIgnoreCase))
		{
			return raw.Substring("about:kkyfb:".Length).Trim('/', ' ');
		}
		return string.Empty;
	}

	private static void ApplyInlinePreviewAction(WebBrowser browser, string action, AuditResult result)
	{
		try
		{
			string encodedPath = action.Substring("preview-inline/".Length);
			string sourcePath = Uri.UnescapeDataString(encodedPath ?? string.Empty);
			string fullPath = Path.GetFullPath(sourcePath);
			if (!File.Exists(fullPath))
			{
				InvokePreviewFailure(browser, "Audit preview PNG does not exist: " + fullPath);
				return;
			}
			FileInfo info = new FileInfo(fullPath);
			if (info.Length <= 0)
			{
				InvokePreviewFailure(browser, "Audit preview PNG is empty: " + fullPath);
				return;
			}
			if (info.Length > 8388608)
			{
				InvokePreviewFailure(browser, "Audit preview PNG is too large for inline display: " + fullPath);
				return;
			}
			string dataUri = "data:image/png;base64," + Convert.ToBase64String(File.ReadAllBytes(fullPath));
			if (browser.Document != null)
			{
				browser.Document.InvokeScript("applyInlinePreviewDataUri", new object[] { dataUri, "Audit inline preview fallback" });
			}
		}
		catch (Exception ex)
		{
			result.Warnings.Add("Audit inline preview handling failed: " + DescribeException(ex));
			InvokePreviewFailure(browser, "Audit inline preview handling failed: " + ex.Message);
		}
	}

	private static void InvokePreviewFailure(WebBrowser browser, string detail)
	{
		try
		{
			if (browser != null && browser.Document != null)
			{
				browser.Document.InvokeScript("showInlinePreviewFailure", new object[] { detail });
			}
		}
		catch
		{
		}
	}

	private static string BodyAttribute(WebBrowser browser, string name)
	{
		try
		{
			return browser.Document != null && browser.Document.Body != null ? browser.Document.Body.GetAttribute(name) ?? string.Empty : string.Empty;
		}
		catch
		{
			return string.Empty;
		}
	}

	private static bool WaitUntil(Func<bool> predicate, int timeoutMs)
	{
		Stopwatch sw = Stopwatch.StartNew();
		while (sw.ElapsedMilliseconds < timeoutMs)
		{
			Application.DoEvents();
			if (predicate())
			{
				return true;
			}
			Thread.Sleep(25);
		}
		return predicate();
	}

	private static void DoEventsFor(int milliseconds)
	{
		Stopwatch sw = Stopwatch.StartNew();
		while (sw.ElapsedMilliseconds < milliseconds)
		{
			Application.DoEvents();
			Thread.Sleep(15);
		}
	}

	private static string Describe(Clickable clickable)
	{
		return clickable.Tag + "#" + clickable.Index + " text='" + clickable.Text + "' href='" + clickable.Href + "' onclick='" + Compact(clickable.OnClick) + "'";
	}

	private static string Compact(string value)
	{
		if (string.IsNullOrWhiteSpace(value))
		{
			return string.Empty;
		}
		string compact = string.Join(" ", value.Split(new[] { '\r', '\n', '\t', ' ' }, StringSplitOptions.RemoveEmptyEntries));
		return compact.Length > 120 ? compact.Substring(0, 117) + "..." : compact;
	}

	private static string BuildConsoleSummary(AuditResult result)
	{
		StringBuilder sb = new StringBuilder();
		sb.AppendLine("Scenario: " + result.Scenario);
		sb.AppendLine("Clickable: " + result.ClickableCount + ", browser clicks: " + result.BrowserClickCount + ", host candidates: " + result.HostActionCandidateCount);
		sb.AppendLine("Performance ms: startup-shell=" + result.StartupShellRenderMilliseconds + ", cold-render=" + result.ColdHtmlRenderMilliseconds + ", warm-render=" + result.HtmlRenderMilliseconds + ", document=" + result.DocumentLoadMilliseconds + ", ready=" + result.DashboardReadyMilliseconds + ", filter=" + result.FilterMilliseconds + ", theme=" + result.ThemeToggleMilliseconds + ", rows=visible " + result.VisibleRowCount + "/DOM " + result.DomRowCount + "/total " + result.DataRowCount);
		sb.AppendLine("Cache ms: save=" + result.CacheSaveMilliseconds + ", cold=" + result.CacheColdLoadMilliseconds + ", warm=" + result.CacheWarmLoadMilliseconds + ", offline=" + result.CacheOfflineLoadMilliseconds + ", bytes=" + result.CacheBytes);
		sb.AppendLine("Failures: " + result.Failures.Count + ", warnings: " + result.Warnings.Count);
		foreach (string failure in result.Failures)
		{
			sb.AppendLine("FAIL " + failure);
		}
		foreach (string warning in result.Warnings.Take(12))
		{
			sb.AppendLine("WARN " + warning);
		}
		return sb.ToString();
	}

	private static string DescribeException(Exception ex)
	{
		if (ex is TargetInvocationException && ex.InnerException != null)
		{
			ex = ex.InnerException;
		}
		return ex.GetType().Name + ": " + ex.Message;
	}

	private static string BuildJson(AuditResult result)
	{
		StringBuilder sb = new StringBuilder();
		sb.AppendLine("{");
		sb.AppendLine("  \"scenario\": " + Json(result.Scenario) + ",");
		sb.AppendLine("  \"hostAssembly\": " + Json(result.HostAssembly) + ",");
		sb.AppendLine("  \"clickableCount\": " + result.ClickableCount + ",");
		sb.AppendLine("  \"browserClickCount\": " + result.BrowserClickCount + ",");
		sb.AppendLine("  \"hostActionCandidateCount\": " + result.HostActionCandidateCount + ",");
		sb.AppendLine("  \"htmlRenderMilliseconds\": " + result.HtmlRenderMilliseconds + ",");
		sb.AppendLine("  \"coldHtmlRenderMilliseconds\": " + result.ColdHtmlRenderMilliseconds + ",");
		sb.AppendLine("  \"startupShellRenderMilliseconds\": " + result.StartupShellRenderMilliseconds + ",");
		sb.AppendLine("  \"startupShellLength\": " + result.StartupShellLength + ",");
		sb.AppendLine("  \"htmlLength\": " + result.HtmlLength + ",");
		sb.AppendLine("  \"documentLoadMilliseconds\": " + result.DocumentLoadMilliseconds + ",");
		sb.AppendLine("  \"dashboardReadyMilliseconds\": " + result.DashboardReadyMilliseconds + ",");
		sb.AppendLine("  \"filterMilliseconds\": " + result.FilterMilliseconds + ",");
		sb.AppendLine("  \"themeToggleMilliseconds\": " + result.ThemeToggleMilliseconds + ",");
		sb.AppendLine("  \"dataRowCount\": " + result.DataRowCount + ",");
		sb.AppendLine("  \"domRowCount\": " + result.DomRowCount + ",");
		sb.AppendLine("  \"visibleRowCount\": " + result.VisibleRowCount + ",");
		sb.AppendLine("  \"cacheBytes\": " + result.CacheBytes + ",");
		sb.AppendLine("  \"cacheSaveMilliseconds\": " + result.CacheSaveMilliseconds + ",");
		sb.AppendLine("  \"cacheColdLoadMilliseconds\": " + result.CacheColdLoadMilliseconds + ",");
		sb.AppendLine("  \"cacheWarmLoadMilliseconds\": " + result.CacheWarmLoadMilliseconds + ",");
		sb.AppendLine("  \"cacheOfflineLoadMilliseconds\": " + result.CacheOfflineLoadMilliseconds + ",");
		sb.AppendLine("  \"hostActions\": " + JsonArray(result.HostActions) + ",");
		sb.AppendLine("  \"warnings\": " + JsonArray(result.Warnings) + ",");
		sb.AppendLine("  \"failures\": " + JsonArray(result.Failures));
		sb.AppendLine("}");
		return sb.ToString();
	}

	private static string JsonArray(IEnumerable<string> values)
	{
		return "[" + string.Join(", ", values.Select(Json).ToArray()) + "]";
	}

	private static string Json(string value)
	{
		if (value == null)
		{
			return "null";
		}
		return "\"" + value.Replace("\\", "\\\\").Replace("\"", "\\\"").Replace("\r", "\\r").Replace("\n", "\\n") + "\"";
	}
}
