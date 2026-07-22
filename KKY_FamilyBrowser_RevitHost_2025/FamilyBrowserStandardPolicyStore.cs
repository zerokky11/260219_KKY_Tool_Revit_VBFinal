using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Security.Cryptography;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.VisualBasic.CompilerServices;

public sealed class FamilyBrowserStandardPolicyStore
{
	private const int ManagedPolicyProbeTimeoutMilliseconds = 800;

	private static readonly object ManagedPolicyProbeSyncRoot = RuntimeHelpers.GetObjectValue(new object());

	private static readonly object PolicyMutationSyncRoot = RuntimeHelpers.GetObjectValue(new object());

	private static string _lastManagedPolicyProbePath = string.Empty;

	private static DateTime _lastManagedPolicyProbeUtc = DateTime.MinValue;

	private static string _lastManagedPolicyProbeResult = string.Empty;

	private FamilyBrowserStandardPolicyStore()
	{
	}

	private static string T(string englishText, string koreanText)
	{
		return FamilyBrowserLanguageService.Text(englishText, koreanText);
	}

	public static string GetPolicyPath(string workspaceRoot)
	{
		string managedPolicyPath = FamilyBrowserMachineConfigStore.ResolveManagedPolicyPath();
		if (!string.IsNullOrWhiteSpace(managedPolicyPath))
		{
			return managedPolicyPath;
		}
		return GetLocalPolicyPath(workspaceRoot);
	}

	public static string GetLocalPolicyPath(string workspaceRoot)
	{
		string folder = GetLocalRegistryFolder(workspaceRoot);
		if (string.IsNullOrWhiteSpace(folder))
		{
			return string.Empty;
		}
		return Path.Combine(folder, "standard-policy.json");
	}

	public static string GetConfiguredManagedPolicyPath()
	{
		return FamilyBrowserMachineConfigStore.ResolveManagedPolicyPath();
	}

	public static bool IsConfiguredManagedPolicyAvailable()
	{
		return !string.IsNullOrWhiteSpace(ResolveUsableManagedPolicyPath());
	}

	public static bool HasUnavailableManagedPolicyPath()
	{
		if (string.IsNullOrWhiteSpace(GetConfiguredManagedPolicyPath()))
		{
			return false;
		}
		return string.IsNullOrWhiteSpace(ResolveUsableManagedPolicyPath());
	}

	public static FamilyBrowserStandardPolicy LoadOrCreate(string workspaceRoot, string currentUser)
	{
		if (string.IsNullOrWhiteSpace(FamilyBrowserMachineConfigStore.ResolveManagedPolicyPath()))
		{
			return CreateDefaultPolicy(currentUser);
		}
		string policyPath = GetPolicyPath(workspaceRoot);
		if (File.Exists(policyPath))
		{
			try
			{
				FamilyBrowserStandardPolicy familyBrowserStandardPolicy = DataContractJsonFileStore.Load<FamilyBrowserStandardPolicy>(policyPath);
				NormalizePolicy(familyBrowserStandardPolicy);
				return familyBrowserStandardPolicy;
			}
			catch (Exception projectError)
			{
				ProjectData.SetProjectError(projectError);
				ProjectData.ClearProjectError();
				throw new InvalidDataException("The managed Family Browser policy exists but could not be read. It was not overwritten.", projectError);
			}
		}
		return WithPolicyMutationLock(workspaceRoot, [SpecialName] () =>
		{
			string latestPath = GetPolicyPath(workspaceRoot);
			if (File.Exists(latestPath))
			{
				FamilyBrowserStandardPolicy latest = DataContractJsonFileStore.Load<FamilyBrowserStandardPolicy>(latestPath);
				NormalizePolicy(latest);
				return latest;
			}
			FamilyBrowserStandardPolicy created = CreateDefaultPolicy(currentUser);
			SaveUnlocked(workspaceRoot, created, currentUser);
			return created;
		});
	}

	public static string Save(string workspaceRoot, FamilyBrowserStandardPolicy policy, string currentUser)
	{
		if (string.IsNullOrWhiteSpace(FamilyBrowserMachineConfigStore.ResolveManagedPolicyPath()))
		{
			throw new InvalidOperationException(BuildManagedDataRootRequiredMessage(T("Save Family Browser policy", "Family Browser 정책 저장")));
		}
		return WithPolicyMutationLock(workspaceRoot, [SpecialName] () => SaveUnlocked(workspaceRoot, policy, currentUser));
	}

	private static string SaveUnlocked(string workspaceRoot, FamilyBrowserStandardPolicy policy, string currentUser)
	{
		NormalizePolicy(policy);
		policy.LastUpdatedUtc = DateTime.UtcNow.ToString("O", CultureInfo.InvariantCulture);
		policy.LastUpdatedBy = currentUser ?? string.Empty;
		string policyPath = GetPolicyPath(workspaceRoot);
		EnsureParentFolder(policyPath);
		string temporaryPath = FamilyBrowserAtomicFileService.CreateSiblingTemporaryPath(policyPath);
		try
		{
			byte[] payload = new UTF8Encoding(false).GetBytes(PlainJsonReportWriter.Serialize(policy));
			using (FileStream stream = new FileStream(temporaryPath, FileMode.CreateNew, FileAccess.Write, FileShare.None))
			{
				stream.Write(payload, 0, payload.Length);
				stream.Flush(true);
			}
			FamilyBrowserAtomicFileService.Promote(temporaryPath, policyPath);
		}
		finally
		{
			try
			{
				if (File.Exists(temporaryPath))
				{
					File.Delete(temporaryPath);
				}
			}
			catch
			{
			}
		}
		return policyPath;
	}

	private static FamilyBrowserStandardPolicy LoadLatestPolicyForMutation(string workspaceRoot, string currentUser)
	{
		if (string.IsNullOrWhiteSpace(FamilyBrowserMachineConfigStore.ResolveManagedPolicyPath()))
		{
			return CreateDefaultPolicy(currentUser);
		}
		string policyPath = GetPolicyPath(workspaceRoot);
		if (File.Exists(policyPath))
		{
			try
			{
				FamilyBrowserStandardPolicy familyBrowserStandardPolicy = DataContractJsonFileStore.Load<FamilyBrowserStandardPolicy>(policyPath);
				NormalizePolicy(familyBrowserStandardPolicy);
				return familyBrowserStandardPolicy;
			}
			catch (Exception projectError)
			{
				ProjectData.SetProjectError(projectError);
				ProjectData.ClearProjectError();
				throw new InvalidDataException("The managed Family Browser policy exists but could not be read for mutation. It was not overwritten.", projectError);
			}
		}
		return CreateDefaultPolicy(currentUser);
	}

	private static string MutateLatestPolicy(string workspaceRoot, string currentUser, Action<FamilyBrowserStandardPolicy> mutator)
	{
		if (mutator == null)
		{
			throw new ArgumentNullException("mutator");
		}
		return WithPolicyMutationLock(workspaceRoot, [SpecialName] () =>
		{
			FamilyBrowserStandardPolicy familyBrowserStandardPolicy = LoadLatestPolicyForMutation(workspaceRoot, currentUser);
			mutator(familyBrowserStandardPolicy);
			return SaveUnlocked(workspaceRoot, familyBrowserStandardPolicy, currentUser);
		});
	}

	public static TResult ReadLatestPolicyWithMutationLock<TResult>(string workspaceRoot, string currentUser, Func<FamilyBrowserStandardPolicy, TResult> reader)
	{
		if (reader == null)
		{
			throw new ArgumentNullException("reader");
		}
		return WithPolicyMutationLock(workspaceRoot, [SpecialName] () =>
		{
			FamilyBrowserStandardPolicy latest = LoadLatestPolicyForMutation(workspaceRoot, currentUser);
			return reader(latest);
		});
	}

	private static TResult WithPolicyMutationLock<TResult>(string workspaceRoot, Func<TResult> action)
	{
		if (action == null)
		{
			throw new ArgumentNullException("action");
		}
		string mutexName = BuildPolicyMutexName(GetPolicyPath(workspaceRoot));
		bool acquired = false;
		using Mutex mutex = new Mutex(initiallyOwned: false, mutexName);
		try
		{
			try
			{
				acquired = mutex.WaitOne(TimeSpan.FromSeconds(30.0));
			}
			catch (AbandonedMutexException ex)
			{
				ProjectData.SetProjectError(ex);
				AbandonedMutexException ex2 = ex;
				acquired = true;
				ProjectData.ClearProjectError();
			}
			if (!acquired)
			{
				throw new IOException("Timed out waiting for the Family Browser standard policy lock.");
			}
			object policyMutationSyncRoot = PolicyMutationSyncRoot;
			ObjectFlowControl.CheckForSyncLockOnValueType(policyMutationSyncRoot);
			bool lockTaken = false;
			try
			{
				Monitor.Enter(policyMutationSyncRoot, ref lockTaken);
				using (FileStream policyFileLease = AcquirePolicyFileMutationLock(GetPolicyPath(workspaceRoot), TimeSpan.FromSeconds(10.0)))
				{
					return action();
				}
			}
			finally
			{
				if (lockTaken)
				{
					Monitor.Exit(policyMutationSyncRoot);
				}
			}
		}
		finally
		{
			if (acquired)
			{
				mutex.ReleaseMutex();
			}
		}
	}

	private static string BuildPolicyMutexName(string policyPath)
	{
		string normalizedPath = (policyPath ?? string.Empty).Trim().ToUpperInvariant();
		using SHA256 sha = SHA256.Create();
		byte[] hashBytes = sha.ComputeHash(Encoding.UTF8.GetBytes(normalizedPath));
		return "KKYFamilyBrowserStandardPolicy-" + BitConverter.ToString(hashBytes).Replace("-", string.Empty).Substring(0, 32);
	}

	public static string ResetToDefault(string workspaceRoot, string currentUser)
	{
		FamilyBrowserStandardPolicy policy = CreateDefaultPolicy(currentUser);
		return Save(workspaceRoot, policy, currentUser);
	}

	public static FamilyBrowserStandardLibrarySlot GetEffectiveSlot(FamilyBrowserStandardPolicy policy)
	{
		NormalizePolicy(policy);
		if (string.Equals(policy.Mode, "Integrated", StringComparison.OrdinalIgnoreCase))
		{
			return policy.IntegratedLibrary;
		}
		string b = FamilyBrowserPolicyKey.Normalize(policy.ActiveDiscipline);
		FamilyBrowserStandardLibrarySlot slot = policy.DisciplineLibraries.FirstOrDefault([SpecialName] (FamilyBrowserStandardLibrarySlot x) => x != null && x.Enabled && string.Equals(FamilyBrowserPolicyKey.Normalize(x.Discipline), b, StringComparison.OrdinalIgnoreCase));
		if (slot == null)
		{
			slot = policy.DisciplineLibraries.FirstOrDefault([SpecialName] (FamilyBrowserStandardLibrarySlot x) => x?.Enabled ?? false);
		}
		if (slot == null)
		{
			slot = policy.DisciplineLibraries.FirstOrDefault([SpecialName] (FamilyBrowserStandardLibrarySlot x) => x != null);
		}
		return slot;
	}

	public static string ResolveEffectiveRegistrationPath(string workspaceRoot, FamilyBrowserStandardPolicy policy)
	{
		FamilyBrowserStandardLibrarySlot slot = GetEffectiveSlot(policy);
		string slotRegistrationPath = ResolveSlotRegistrationPath(workspaceRoot, slot);
		if (!string.IsNullOrWhiteSpace(slotRegistrationPath))
		{
			return slotRegistrationPath;
		}
		string registryFolder = GetRegistryFolder(workspaceRoot);
		if (string.IsNullOrWhiteSpace(registryFolder))
		{
			return string.Empty;
		}
		if (!HasAnyPolicyRegistration(policy))
		{
			return Path.Combine(registryFolder, "active-standard-library.json");
		}
		return Path.Combine(registryFolder, "missing-standard-library-slot.json");
	}

	public static string ResolveSlotRegistrationPath(string workspaceRoot, FamilyBrowserStandardLibrarySlot slot)
	{
		if (slot == null)
		{
			return string.Empty;
		}
		string registryFolder = GetRegistryFolder(workspaceRoot);
		if (string.IsNullOrWhiteSpace(registryFolder))
		{
			return string.Empty;
		}
		string configuredPath = (slot.RegistrationPath ?? string.Empty).Trim();
		if (IsCurrentRevitVersionDataPath(workspaceRoot, configuredPath))
		{
			return configuredPath;
		}
		string safeSlotKey = FamilyBrowserPolicyKey.Normalize(slot.SlotKey);
		if (string.IsNullOrWhiteSpace(safeSlotKey))
		{
			safeSlotKey = FamilyBrowserPolicyKey.Normalize(slot.Discipline);
		}
		if (string.IsNullOrWhiteSpace(safeSlotKey))
		{
			safeSlotKey = "standard";
		}
		string resolvedSlotPath = Path.Combine(registryFolder, "standard-library-" + safeSlotKey + ".json");
		return resolvedSlotPath;
	}

	private static FileStream AcquirePolicyFileMutationLock(string policyPath, TimeSpan timeout)
	{
		if (string.IsNullOrWhiteSpace(policyPath))
		{
			throw new IOException("The Family Browser policy path is empty.");
		}
		EnsureParentFolder(policyPath);
		string lockPath = policyPath + ".kky-lock";
		DateTime deadline = DateTime.UtcNow.Add(timeout <= TimeSpan.Zero ? TimeSpan.FromSeconds(10.0) : timeout);
		Exception lastError = null;
		while (DateTime.UtcNow < deadline)
		{
			try
			{
				return new FileStream(lockPath, FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.None, 1, FileOptions.WriteThrough);
			}
			catch (IOException ex)
			{
				lastError = ex;
			}
			catch (UnauthorizedAccessException ex)
			{
				lastError = ex;
			}
			Thread.Sleep(75);
		}
		throw new IOException("Timed out waiting for the shared Family Browser policy file lock.", lastError);
	}

	public static string ResolveSlotSnapshotPath(string workspaceRoot, FamilyBrowserStandardLibrarySlot slot, StandardLibraryRegistrationRecord registration)
	{
		string registrationPath = ((registration == null) ? string.Empty : registration.LastSnapshotPath);
		if (IsCurrentRevitVersionDataPath(workspaceRoot, registrationPath))
		{
			return registrationPath;
		}
		string slotPath = ((slot == null) ? string.Empty : slot.SnapshotPath);
		if (IsCurrentRevitVersionDataPath(workspaceRoot, slotPath))
		{
			return slotPath;
		}
		return string.Empty;
	}

	public static string GetRegistryFolder(string workspaceRoot)
	{
		if (!string.IsNullOrWhiteSpace(FamilyBrowserMachineConfigStore.ResolveManagedPolicyPath()))
		{
			string versionRoot = GetVersionedManagedDataRootFolder(workspaceRoot);
			if (string.IsNullOrWhiteSpace(versionRoot))
			{
				return string.Empty;
			}
			return Path.Combine(versionRoot, "Registry");
		}
		return GetLocalRegistryFolder(workspaceRoot);
	}

	public static string GetSnapshotFolder(string workspaceRoot)
	{
		if (!string.IsNullOrWhiteSpace(FamilyBrowserMachineConfigStore.ResolveManagedPolicyPath()))
		{
			string versionRoot = GetVersionedManagedDataRootFolder(workspaceRoot);
			if (string.IsNullOrWhiteSpace(versionRoot))
			{
				return string.Empty;
			}
			return Path.Combine(versionRoot, "Snapshots");
		}
		return GetUnavailableManagedDataFolder(workspaceRoot, "Snapshots");
	}

	public static string GetThumbnailFolder(string workspaceRoot)
	{
		if (!string.IsNullOrWhiteSpace(FamilyBrowserMachineConfigStore.ResolveManagedPolicyPath()))
		{
			string versionRoot = GetVersionedManagedDataRootFolder(workspaceRoot);
			if (string.IsNullOrWhiteSpace(versionRoot))
			{
				return string.Empty;
			}
			return Path.Combine(versionRoot, "Thumbnails");
		}
		return GetUnavailableManagedDataFolder(workspaceRoot, "Thumbnails");
	}

	public static string GetStandardListFolder(string workspaceRoot)
	{
		if (!string.IsNullOrWhiteSpace(FamilyBrowserMachineConfigStore.ResolveManagedPolicyPath()))
		{
			string dataRoot = GetDataRootFolder(workspaceRoot);
			if (string.IsNullOrWhiteSpace(dataRoot))
			{
				return string.Empty;
			}
			return Path.Combine(dataRoot, "StandardLists");
		}
		return GetUnavailableManagedDataFolder(workspaceRoot, "StandardLists");
	}

	public static string GetDataRootFolder(string workspaceRoot)
	{
		string managedPolicyPath = FamilyBrowserMachineConfigStore.ResolveManagedPolicyPath();
		if (!string.IsNullOrWhiteSpace(managedPolicyPath))
		{
			string rootFolder = ResolveManagedDataRoot(ResolveManagedPolicyFolder(managedPolicyPath, workspaceRoot));
			if (!string.IsNullOrWhiteSpace(rootFolder))
			{
				return rootFolder;
			}
		}
		return GetUnavailableManagedDataFolder(workspaceRoot, string.Empty);
	}

	public static string GetVersionedManagedDataRootFolder(string workspaceRoot)
	{
		string dataRoot = GetDataRootFolder(workspaceRoot);
		if (string.IsNullOrWhiteSpace(dataRoot))
		{
			return string.Empty;
		}
		return FamilyBrowserRevitVersionContext.VersionedDataRoot(dataRoot);
	}

	public static bool IsCurrentRevitVersionDataPath(string workspaceRoot, string path)
	{
		if (string.IsNullOrWhiteSpace(path))
		{
			return false;
		}
		string dataRoot = GetDataRootFolder(workspaceRoot);
		if (string.IsNullOrWhiteSpace(dataRoot))
		{
			return false;
		}
		return FamilyBrowserRevitVersionContext.IsPathInCurrentVersionRoot(path, dataRoot);
	}

	public static string GetDataFolder(string workspaceRoot, string folderName)
	{
		string safeFolderName = (folderName ?? string.Empty).Trim();
		if (string.IsNullOrWhiteSpace(safeFolderName))
		{
			return GetDataRootFolder(workspaceRoot);
		}
		string dataRoot = GetDataRootFolder(workspaceRoot);
		if (string.IsNullOrWhiteSpace(dataRoot))
		{
			return string.Empty;
		}
		if (string.Equals(safeFolderName, "Projects", StringComparison.OrdinalIgnoreCase))
		{
			string versionRoot = FamilyBrowserRevitVersionContext.VersionedDataRoot(dataRoot);
			if (string.IsNullOrWhiteSpace(versionRoot))
			{
				return string.Empty;
			}
			return Path.Combine(versionRoot, safeFolderName);
		}
		return Path.Combine(dataRoot, safeFolderName);
	}

	public static bool IsManagedDataRootAvailable(string workspaceRoot = "")
	{
		return !string.IsNullOrWhiteSpace(ResolveManagedDataRootForWrite(workspaceRoot, throwIfUnavailable: false));
	}

	public static string ResolveManagedDataRootForWrite(string workspaceRoot, bool throwIfUnavailable = true, string operationName = "")
	{
		string managedPolicyPath = FamilyBrowserMachineConfigStore.ResolveManagedPolicyPath();
		if (!string.IsNullOrWhiteSpace(managedPolicyPath) && IsManagedPolicyPathUsableFast(managedPolicyPath))
		{
			string rootFolder = ResolveManagedDataRoot(ResolveManagedPolicyFolder(managedPolicyPath, workspaceRoot));
			if (!string.IsNullOrWhiteSpace(rootFolder))
			{
				return rootFolder;
			}
		}
		if (throwIfUnavailable)
		{
			throw new InvalidOperationException(BuildManagedDataRootRequiredMessage(operationName));
		}
		return string.Empty;
	}

	public static void RequireManagedDataRootForWrite(string workspaceRoot, string operationName)
	{
		ResolveManagedDataRootForWrite(workspaceRoot, throwIfUnavailable: true, operationName);
	}

	private static string BuildManagedDataRootRequiredMessage(string operationName)
	{
		string caption = (operationName ?? string.Empty).Trim();
		if (string.IsNullOrWhiteSpace(caption))
		{
			caption = T("Family Browser data write", "Family Browser 데이터 저장");
		}
		return caption + ": " + T("Family Browser scan data is not written to the local C fallback folder. Refresh the homepage path and connect a managed shared folder before running scans, comparisons, thumbnails, or apply reports.", "Family Browser 스캔 데이터는 로컬 C fallback 폴더에 저장하지 않습니다. 홈페이지 경로를 다시 확인해서 공용 관리 폴더를 연결한 뒤 스캔, 비교, 3D 이미지, 적용 리포트를 실행하세요.");
	}

	public static string AttachRegistrationToEffectiveSlot(string workspaceRoot, FamilyBrowserStandardPolicy policy, StandardLibraryRegistrationRecord registration, string currentUser)
	{
		if (registration == null)
		{
			throw new ArgumentNullException("registration");
		}
		NormalizePolicy(policy);
		FamilyBrowserStandardLibrarySlot requestedSlot = GetEffectiveSlot(policy);
		if (requestedSlot == null)
		{
			throw new InvalidOperationException(T("No effective standard library slot is available.", "사용 가능한 표준 라이브러리 슬롯이 없습니다."));
		}
		return AttachRegistrationToSlot(workspaceRoot, policy, requestedSlot.SlotKey, registration, currentUser);
	}

	public static string AttachRegistrationToSlot(string workspaceRoot, FamilyBrowserStandardPolicy policy, string slotKey, StandardLibraryRegistrationRecord registration, string currentUser)
	{
		if (registration == null)
		{
			throw new ArgumentNullException("registration");
		}
		NormalizePolicy(policy);
		FamilyBrowserStandardLibrarySlot requestedSlot = FindSlotByKey(policy, slotKey);
		if (requestedSlot == null)
		{
			requestedSlot = GetEffectiveSlot(policy);
		}
		if (requestedSlot == null)
		{
			throw new InvalidOperationException(T("No standard library slot is available.", "표준 라이브러리 슬롯이 없습니다."));
		}
		FamilyBrowserStandardLibrarySlot familyBrowserStandardLibrarySlot = CloneSlot(requestedSlot);
		string slotKey2 = familyBrowserStandardLibrarySlot.SlotKey ?? string.Empty;
		return WithPolicyMutationLock(workspaceRoot, [SpecialName] () =>
		{
			FamilyBrowserStandardPolicy policy2 = LoadLatestPolicyForMutation(workspaceRoot, currentUser);
			FamilyBrowserStandardLibrarySlot familyBrowserStandardLibrarySlot2 = ResolveMutableSlot(policy2, slotKey2, familyBrowserStandardLibrarySlot);
			if (familyBrowserStandardLibrarySlot2 == null)
			{
				throw new InvalidOperationException(T("No standard library slot is available.", "표준 라이브러리 슬롯이 없습니다."));
			}
			string text = SaveSlotRegistration(workspaceRoot, familyBrowserStandardLibrarySlot2.SlotKey, registration);
			ApplyRegistrationToSlot(familyBrowserStandardLibrarySlot2, text, registration);
			SaveUnlocked(workspaceRoot, policy2, currentUser);
			return text;
		});
	}

	public static void SetStandardListForSlot(string workspaceRoot, FamilyBrowserStandardPolicy policy, string slotKey, string standardListPath, string sheetName, string currentUser)
	{
		NormalizePolicy(policy);
		FamilyBrowserStandardLibrarySlot requestedSlot = FindSlotByKey(policy, slotKey);
		if (requestedSlot == null)
		{
			requestedSlot = GetEffectiveSlot(policy);
		}
		if (requestedSlot == null)
		{
			throw new InvalidOperationException(T("No standard library slot is available.", "표준 라이브러리 슬롯이 없습니다."));
		}
		FamilyBrowserStandardLibrarySlot familyBrowserStandardLibrarySlot = CloneSlot(requestedSlot);
		string slotKey2 = familyBrowserStandardLibrarySlot.SlotKey ?? string.Empty;
		string standardListPath2 = (standardListPath ?? string.Empty).Trim();
		string standardListSheetName = (sheetName ?? string.Empty).Trim();
		MutateLatestPolicy(workspaceRoot, currentUser, [SpecialName] (FamilyBrowserStandardPolicy latest) =>
		{
			FamilyBrowserStandardLibrarySlot familyBrowserStandardLibrarySlot2 = ResolveMutableSlot(latest, slotKey2, familyBrowserStandardLibrarySlot);
			if (familyBrowserStandardLibrarySlot2 != null)
			{
				familyBrowserStandardLibrarySlot2.StandardListPath = standardListPath2;
				familyBrowserStandardLibrarySlot2.StandardListSheetName = standardListSheetName;
			}
			FamilyBrowserStandardLibrarySlot familyBrowserStandardLibrarySlot3 = ResolveMutableSlot(policy, slotKey2, familyBrowserStandardLibrarySlot);
			if (familyBrowserStandardLibrarySlot3 != null)
			{
				familyBrowserStandardLibrarySlot3.StandardListPath = standardListPath2;
				familyBrowserStandardLibrarySlot3.StandardListSheetName = standardListSheetName;
			}
		});
	}

	public static void ClearRegistrationForSlot(string workspaceRoot, FamilyBrowserStandardPolicy policy, string slotKey, string currentUser)
	{
		NormalizePolicy(policy);
		FamilyBrowserStandardLibrarySlot requestedSlot = FindSlotByKey(policy, slotKey);
		if (requestedSlot == null)
		{
			requestedSlot = GetEffectiveSlot(policy);
		}
		if (requestedSlot == null)
		{
			throw new InvalidOperationException(T("No standard library slot is available.", "표준 라이브러리 슬롯이 없습니다."));
		}
		FamilyBrowserStandardLibrarySlot familyBrowserStandardLibrarySlot = CloneSlot(requestedSlot);
		string slotKey2 = familyBrowserStandardLibrarySlot.SlotKey ?? string.Empty;
		MutateLatestPolicy(workspaceRoot, currentUser, [SpecialName] (FamilyBrowserStandardPolicy latest) =>
		{
			ClearRegistrationFields(ResolveMutableSlot(latest, slotKey2, familyBrowserStandardLibrarySlot));
			ClearRegistrationFields(ResolveMutableSlot(policy, slotKey2, familyBrowserStandardLibrarySlot));
		});
	}

	private static FamilyBrowserStandardLibrarySlot FindSlotByKey(FamilyBrowserStandardPolicy policy, string slotKey)
	{
		if (policy == null)
		{
			return null;
		}
		string text = FamilyBrowserPolicyKey.Normalize(slotKey);
		if (string.IsNullOrWhiteSpace(text))
		{
			return null;
		}
		if (policy.IntegratedLibrary != null && string.Equals(FamilyBrowserPolicyKey.Normalize(policy.IntegratedLibrary.SlotKey), text, StringComparison.OrdinalIgnoreCase))
		{
			return policy.IntegratedLibrary;
		}
		if (policy.DisciplineLibraries == null)
		{
			return null;
		}
		return policy.DisciplineLibraries.FirstOrDefault([SpecialName] (FamilyBrowserStandardLibrarySlot x) => x != null && string.Equals(FamilyBrowserPolicyKey.Normalize(x.SlotKey), text, StringComparison.OrdinalIgnoreCase));
	}

	private static FamilyBrowserStandardLibrarySlot ResolveMutableSlot(FamilyBrowserStandardPolicy policy, string slotKey, FamilyBrowserStandardLibrarySlot requestedSlot)
	{
		NormalizePolicy(policy);
		FamilyBrowserStandardLibrarySlot slot = FindSlotByKey(policy, slotKey);
		if (slot != null)
		{
			return slot;
		}
		slot = FindEquivalentSlot(policy, requestedSlot);
		if (slot != null)
		{
			return slot;
		}
		if (requestedSlot == null)
		{
			return GetEffectiveSlot(policy);
		}
		FamilyBrowserStandardLibrarySlot clonedSlot = CloneSlot(requestedSlot);
		NormalizeSlot(clonedSlot, "discipline-" + FamilyBrowserPolicyKey.Normalize(clonedSlot.Discipline), clonedSlot.Discipline, clonedSlot.DisplayName);
		if (string.Equals(FamilyBrowserPolicyKey.Normalize(clonedSlot.SlotKey), "integrated", StringComparison.OrdinalIgnoreCase) || string.Equals(FamilyBrowserPolicyKey.Normalize(clonedSlot.Discipline), FamilyBrowserPolicyKey.Normalize("Integrated"), StringComparison.OrdinalIgnoreCase))
		{
			policy.IntegratedLibrary = clonedSlot;
			return policy.IntegratedLibrary;
		}
		if (policy.DisciplineLibraries == null)
		{
			policy.DisciplineLibraries = new List<FamilyBrowserStandardLibrarySlot>();
		}
		policy.DisciplineLibraries.Add(clonedSlot);
		return clonedSlot;
	}

	private static FamilyBrowserStandardLibrarySlot FindEquivalentSlot(FamilyBrowserStandardPolicy policy, FamilyBrowserStandardLibrarySlot requestedSlot)
	{
		if (policy == null || requestedSlot == null)
		{
			return null;
		}
		string text = FamilyBrowserPolicyKey.Normalize(requestedSlot.Discipline);
		string text2 = FamilyBrowserPolicyKey.Normalize(requestedSlot.DisplayName);
		if (policy.IntegratedLibrary != null)
		{
			if (!string.IsNullOrWhiteSpace(text) && string.Equals(FamilyBrowserPolicyKey.Normalize(policy.IntegratedLibrary.Discipline), text, StringComparison.OrdinalIgnoreCase))
			{
				return policy.IntegratedLibrary;
			}
			if (!string.IsNullOrWhiteSpace(text2) && string.Equals(FamilyBrowserPolicyKey.Normalize(policy.IntegratedLibrary.DisplayName), text2, StringComparison.OrdinalIgnoreCase))
			{
				return policy.IntegratedLibrary;
			}
		}
		if (policy.DisciplineLibraries == null)
		{
			return null;
		}
		return policy.DisciplineLibraries.FirstOrDefault([SpecialName] (FamilyBrowserStandardLibrarySlot x) =>
		{
			if (x == null)
			{
				return false;
			}
			if (!string.IsNullOrWhiteSpace(text) && string.Equals(FamilyBrowserPolicyKey.Normalize(x.Discipline), text, StringComparison.OrdinalIgnoreCase))
			{
				return true;
			}
			return (!string.IsNullOrWhiteSpace(text2) && string.Equals(FamilyBrowserPolicyKey.Normalize(x.DisplayName), text2, StringComparison.OrdinalIgnoreCase)) ? true : false;
		});
	}

	private static FamilyBrowserStandardLibrarySlot CloneSlot(FamilyBrowserStandardLibrarySlot slot)
	{
		if (slot == null)
		{
			return null;
		}
		return new FamilyBrowserStandardLibrarySlot
		{
			SlotKey = (slot.SlotKey ?? string.Empty),
			Discipline = (slot.Discipline ?? string.Empty),
			DisplayName = (slot.DisplayName ?? string.Empty),
			RegistrationPath = (slot.RegistrationPath ?? string.Empty),
			SourceId = (slot.SourceId ?? string.Empty),
			StandardRvtPath = (slot.StandardRvtPath ?? string.Empty),
			SnapshotPath = (slot.SnapshotPath ?? string.Empty),
			StandardListPath = (slot.StandardListPath ?? string.Empty),
			StandardListSheetName = (slot.StandardListSheetName ?? string.Empty),
			LastSnapshotAtUtc = (slot.LastSnapshotAtUtc ?? string.Empty),
			Enabled = slot.Enabled
		};
	}

	private static void ApplyRegistrationToSlot(FamilyBrowserStandardLibrarySlot slot, string slotRegistrationPath, StandardLibraryRegistrationRecord registration)
	{
		if (slot != null && registration != null)
		{
			slot.RegistrationPath = slotRegistrationPath ?? string.Empty;
			slot.SourceId = registration.SourceId;
			slot.StandardRvtPath = registration.ResolvedPath;
			slot.SnapshotPath = registration.LastSnapshotPath;
			slot.LastSnapshotAtUtc = registration.LastSnapshotAtUtc;
			if (string.IsNullOrWhiteSpace(slot.DisplayName))
			{
				slot.DisplayName = (string.IsNullOrWhiteSpace(registration.DisplayName) ? slot.Discipline : registration.DisplayName);
			}
		}
	}

	private static void ClearRegistrationFields(FamilyBrowserStandardLibrarySlot slot)
	{
		if (slot != null)
		{
			slot.RegistrationPath = string.Empty;
			slot.SourceId = string.Empty;
			slot.StandardRvtPath = string.Empty;
			slot.SnapshotPath = string.Empty;
			slot.LastSnapshotAtUtc = string.Empty;
		}
	}

	public static void SetMode(string workspaceRoot, FamilyBrowserStandardPolicy policy, string mode, string currentUser)
	{
		NormalizePolicy(policy);
		string mode2 = (string.Equals(mode, "Integrated", StringComparison.OrdinalIgnoreCase) ? "Integrated" : "DisciplineSeparated");
		MutateLatestPolicy(workspaceRoot, currentUser, [SpecialName] (FamilyBrowserStandardPolicy latest) =>
		{
			latest.Mode = mode2;
			policy.Mode = mode2;
		});
	}

	public static bool IsDetailedSystemTypeComparisonEnabled(FamilyBrowserStandardPolicy policy)
	{
		return policy == null || policy.CompareDetailedSystemTypeComponents != false;
	}

	public static void SetDetailedSystemTypeComparison(string workspaceRoot, FamilyBrowserStandardPolicy policy, bool enabled, string currentUser)
	{
		NormalizePolicy(policy);
		MutateLatestPolicy(workspaceRoot, currentUser, [SpecialName] (FamilyBrowserStandardPolicy latest) =>
		{
			latest.CompareDetailedSystemTypeComponents = enabled;
			policy.CompareDetailedSystemTypeComponents = enabled;
		});
	}

	public static bool IsProjectElementChangeTrackingEnabled(FamilyBrowserStandardPolicy policy)
	{
		if (policy?.FileGuard == null || !policy.FileGuard.Enabled)
		{
			return false;
		}
		return (policy.FileGuard.Targets ?? new List<FamilyBrowserFileGuardTarget>())
			.Any([SpecialName] (FamilyBrowserFileGuardTarget target) => target != null && target.Enabled && target.TrackElementChanges);
	}

	public static void SetActiveDiscipline(string workspaceRoot, FamilyBrowserStandardPolicy policy, string discipline, string currentUser)
	{
		NormalizePolicy(policy);
		string normalizedDiscipline = ResolveDisciplineKey(policy, discipline);
		if (string.IsNullOrWhiteSpace(normalizedDiscipline))
		{
			normalizedDiscipline = "Mechanical";
		}
		string text = normalizedDiscipline;
		MutateLatestPolicy(workspaceRoot, currentUser, [SpecialName] (FamilyBrowserStandardPolicy latest) =>
		{
			string text2 = ResolveDisciplineKey(latest, text);
			if (string.IsNullOrWhiteSpace(text2))
			{
				text2 = text;
			}
			latest.ActiveDiscipline = text2;
			policy.ActiveDiscipline = text2;
		});
	}

	public static List<FamilyBrowserStandardLibrarySlot> GetDisciplineSlots(FamilyBrowserStandardPolicy policy)
	{
		NormalizePolicy(policy);
		return policy.DisciplineLibraries.Where([SpecialName] (FamilyBrowserStandardLibrarySlot x) => x?.Enabled ?? false).OrderBy<FamilyBrowserStandardLibrarySlot, string>([SpecialName] (FamilyBrowserStandardLibrarySlot x) => ResolveSlotDisplayName(x, korean: false), StringComparer.OrdinalIgnoreCase).ToList();
	}

	public static string ResolveDisciplineKey(FamilyBrowserStandardPolicy policy, string value)
	{
		NormalizePolicy(policy);
		string text = (value ?? string.Empty).Trim();
		if (string.IsNullOrWhiteSpace(text))
		{
			return string.Empty;
		}
		string builtIn = FamilyBrowserStandardPolicyStore.ResolveDisciplineKey(text);
		if (!string.IsNullOrWhiteSpace(builtIn))
		{
			return builtIn;
		}
		string text2 = FamilyBrowserPolicyKey.Normalize(text);
		FamilyBrowserStandardLibrarySlot slot = policy.DisciplineLibraries.FirstOrDefault([SpecialName] (FamilyBrowserStandardLibrarySlot x) =>
		{
			if (x == null)
			{
				return false;
			}
			string a = FamilyBrowserPolicyKey.Normalize(x.SlotKey);
			return string.Equals(FamilyBrowserPolicyKey.Normalize(x.Discipline), text2, StringComparison.OrdinalIgnoreCase) || string.Equals(FamilyBrowserPolicyKey.Normalize(x.DisplayName), text2, StringComparison.OrdinalIgnoreCase) || string.Equals(a, text2, StringComparison.OrdinalIgnoreCase) || string.Equals(a, "discipline-" + text2, StringComparison.OrdinalIgnoreCase);
		});
		if (slot == null)
		{
			return text;
		}
		return slot.Discipline;
	}

	public static string ResolveSlotDisplayName(FamilyBrowserStandardLibrarySlot slot, bool korean)
	{
		if (slot == null)
		{
			return ResolveDisciplineLabel("Other", korean);
		}
		string disciplineLabel = ResolveDisciplineLabel(slot.Discipline, korean);
		string displayName = (slot.DisplayName ?? string.Empty).Trim();
		if (string.IsNullOrWhiteSpace(displayName))
		{
			return disciplineLabel;
		}
		if (korean)
		{
			string normalizedDisplay = FamilyBrowserPolicyKey.Normalize(displayName);
			if (string.Equals(normalizedDisplay, FamilyBrowserPolicyKey.Normalize("Architecture"), StringComparison.OrdinalIgnoreCase) || string.Equals(normalizedDisplay, FamilyBrowserPolicyKey.Normalize("Structure"), StringComparison.OrdinalIgnoreCase) || string.Equals(normalizedDisplay, FamilyBrowserPolicyKey.Normalize("Mechanical"), StringComparison.OrdinalIgnoreCase) || string.Equals(normalizedDisplay, FamilyBrowserPolicyKey.Normalize("Electrical"), StringComparison.OrdinalIgnoreCase) || string.Equals(normalizedDisplay, FamilyBrowserPolicyKey.Normalize("Fire Protection"), StringComparison.OrdinalIgnoreCase) || string.Equals(normalizedDisplay, FamilyBrowserPolicyKey.Normalize("Integrated Standard RVT"), StringComparison.OrdinalIgnoreCase) || string.Equals(normalizedDisplay, FamilyBrowserPolicyKey.Normalize("Other"), StringComparison.OrdinalIgnoreCase))
			{
				return disciplineLabel;
			}
		}
		return displayName;
	}

	public static void AddOrUpdateDisciplineSlot(string workspaceRoot, FamilyBrowserStandardPolicy policy, string disciplineName, string displayName, string currentUser)
	{
		NormalizePolicy(policy);
		string text = (displayName ?? string.Empty).Trim();
		if (string.IsNullOrWhiteSpace(text))
		{
			throw new ArgumentException(T("Trade display name is required.", "공종 표시 이름이 필요합니다."), "displayName");
		}
		string text2 = (string.IsNullOrWhiteSpace(disciplineName) ? text : disciplineName.Trim());
		string builtIn = ResolveDisciplineKey(text2);
		if (!string.IsNullOrWhiteSpace(builtIn))
		{
			text2 = builtIn;
		}
		MutateLatestPolicy(workspaceRoot, currentUser, [SpecialName] (FamilyBrowserStandardPolicy latest) =>
		{
			AddOrUpdateDisciplineSlotInPolicy(latest, text2, text);
			AddOrUpdateDisciplineSlotInPolicy(policy, text2, text);
		});
	}

	public static void RenameDisciplineSlot(string workspaceRoot, FamilyBrowserStandardPolicy policy, string disciplineName, string displayName, string currentUser)
	{
		NormalizePolicy(policy);
		string text = (displayName ?? string.Empty).Trim();
		if (string.IsNullOrWhiteSpace(text))
		{
			throw new ArgumentException(T("Trade display name is required.", "공종 표시 이름이 필요합니다."), "displayName");
		}
		string disciplineName2 = ResolveDisciplineKey(policy, disciplineName);
		MutateLatestPolicy(workspaceRoot, currentUser, [SpecialName] (FamilyBrowserStandardPolicy latest) =>
		{
			RenameDisciplineSlotInPolicy(latest, disciplineName2, text);
			RenameDisciplineSlotInPolicy(policy, disciplineName2, text);
		});
	}

	public static void RemoveDisciplineSlot(string workspaceRoot, FamilyBrowserStandardPolicy policy, string disciplineName, string currentUser)
	{
		NormalizePolicy(policy);
		string disciplineName2 = ResolveDisciplineKey(policy, disciplineName);
		MutateLatestPolicy(workspaceRoot, currentUser, [SpecialName] (FamilyBrowserStandardPolicy latest) =>
		{
			RemoveDisciplineSlotInPolicy(latest, disciplineName2);
			RemoveDisciplineSlotInPolicy(policy, disciplineName2);
		});
	}

	private static void AddOrUpdateDisciplineSlotInPolicy(FamilyBrowserStandardPolicy policy, string disciplineName, string displayName)
	{
		NormalizePolicy(policy);
		string display = (displayName ?? string.Empty).Trim();
		string discipline = (string.IsNullOrWhiteSpace(disciplineName) ? display : disciplineName.Trim());
		string builtIn = ResolveDisciplineKey(discipline);
		if (!string.IsNullOrWhiteSpace(builtIn))
		{
			discipline = builtIn;
		}
		string b = FamilyBrowserPolicyKey.Normalize(discipline);
		string b2 = FamilyBrowserPolicyKey.Normalize(display);
		FamilyBrowserStandardLibrarySlot slot = policy.DisciplineLibraries.FirstOrDefault([SpecialName] (FamilyBrowserStandardLibrarySlot x) => x != null && (string.Equals(FamilyBrowserPolicyKey.Normalize(x.Discipline), b, StringComparison.OrdinalIgnoreCase) || string.Equals(FamilyBrowserPolicyKey.Normalize(x.DisplayName), b2, StringComparison.OrdinalIgnoreCase)));
		if (slot == null)
		{
			slot = FamilyBrowserStandardLibrarySlot.CreateDiscipline(discipline, display);
			policy.DisciplineLibraries.Add(slot);
		}
		slot.DisplayName = display;
		slot.Enabled = true;
		if (string.IsNullOrWhiteSpace(slot.SlotKey))
		{
			slot.SlotKey = "discipline-" + FamilyBrowserPolicyKey.Normalize(slot.Discipline);
		}
		if (string.IsNullOrWhiteSpace(policy.ActiveDiscipline))
		{
			policy.ActiveDiscipline = slot.Discipline;
		}
	}

	private static void RenameDisciplineSlotInPolicy(FamilyBrowserStandardPolicy policy, string disciplineName, string displayName)
	{
		NormalizePolicy(policy);
		string discipline = ResolveDisciplineKey(policy, disciplineName);
		string b = FamilyBrowserPolicyKey.Normalize(discipline);
		(policy.DisciplineLibraries.FirstOrDefault([SpecialName] (FamilyBrowserStandardLibrarySlot x) => x != null && string.Equals(FamilyBrowserPolicyKey.Normalize(x.Discipline), b, StringComparison.OrdinalIgnoreCase)) ?? throw new InvalidOperationException(T("Active trade slot was not found.", "활성 공종 슬롯을 찾지 못했습니다."))).DisplayName = displayName;
	}

	private static void RemoveDisciplineSlotInPolicy(FamilyBrowserStandardPolicy policy, string disciplineName)
	{
		NormalizePolicy(policy);
		string discipline = ResolveDisciplineKey(policy, disciplineName);
		string b = FamilyBrowserPolicyKey.Normalize(discipline);
		FamilyBrowserStandardLibrarySlot? obj = policy.DisciplineLibraries.FirstOrDefault([SpecialName] (FamilyBrowserStandardLibrarySlot x) => x != null && string.Equals(FamilyBrowserPolicyKey.Normalize(x.Discipline), b, StringComparison.OrdinalIgnoreCase)) ?? throw new InvalidOperationException(T("Active trade slot was not found.", "활성 공종 슬롯을 찾지 못했습니다."));
		int enabledCount = policy.DisciplineLibraries.Where([SpecialName] (FamilyBrowserStandardLibrarySlot x) => x?.Enabled ?? false).Count();
		if (obj.Enabled && enabledCount <= 1)
		{
			throw new InvalidOperationException(T("At least one trade target must remain.", "공종 대상은 최소 하나 이상 남아 있어야 합니다."));
		}
		obj.Enabled = false;
		obj.RegistrationPath = string.Empty;
		obj.SourceId = string.Empty;
		obj.StandardRvtPath = string.Empty;
		obj.SnapshotPath = string.Empty;
		obj.StandardListPath = string.Empty;
		obj.StandardListSheetName = string.Empty;
		obj.LastSnapshotAtUtc = string.Empty;
		if (string.Equals(FamilyBrowserPolicyKey.Normalize(policy.ActiveDiscipline), b, StringComparison.OrdinalIgnoreCase))
		{
			FamilyBrowserStandardLibrarySlot nextSlot = policy.DisciplineLibraries.FirstOrDefault([SpecialName] (FamilyBrowserStandardLibrarySlot x) => x?.Enabled ?? false);
			if (nextSlot != null)
			{
				policy.ActiveDiscipline = nextSlot.Discipline;
			}
			else
			{
				policy.ActiveDiscipline = "Mechanical";
			}
		}
	}

	public static string ResolveDisciplineLabel(string discipline, bool korean)
	{
		string left = FamilyBrowserPolicyKey.Normalize(discipline);
		if (Operators.CompareString(left, FamilyBrowserPolicyKey.Normalize("Architecture"), TextCompare: false) == 0)
		{
			return korean ? "건축" : "Architecture";
		}
		if (Operators.CompareString(left, FamilyBrowserPolicyKey.Normalize("Structure"), TextCompare: false) == 0)
		{
			return korean ? "구조" : "Structure";
		}
		if (Operators.CompareString(left, FamilyBrowserPolicyKey.Normalize("Mechanical"), TextCompare: false) == 0)
		{
			return korean ? "설비" : "Mechanical";
		}
		if (Operators.CompareString(left, FamilyBrowserPolicyKey.Normalize("Electrical"), TextCompare: false) == 0)
		{
			return korean ? "전기" : "Electrical";
		}
		if (Operators.CompareString(left, FamilyBrowserPolicyKey.Normalize("FireProtection"), TextCompare: false) == 0)
		{
			return korean ? "소방" : "Fire";
		}
		if (Operators.CompareString(left, FamilyBrowserPolicyKey.Normalize("Integrated"), TextCompare: false) == 0)
		{
			return korean ? "통합" : "Integrated";
		}
		if (string.IsNullOrWhiteSpace(discipline))
		{
			return korean ? "기타" : "Other";
		}
		return discipline;
	}

	public static string ResolveDisciplineKey(string value)
	{
		string normalizedValue = FamilyBrowserPolicyKey.Normalize(value);
		if (string.Equals(normalizedValue, "건축", StringComparison.OrdinalIgnoreCase))
		{
			return "Architecture";
		}
		if (string.Equals(normalizedValue, "구조", StringComparison.OrdinalIgnoreCase))
		{
			return "Structure";
		}
		if (string.Equals(normalizedValue, "설비", StringComparison.OrdinalIgnoreCase) || string.Equals(normalizedValue, "기계", StringComparison.OrdinalIgnoreCase))
		{
			return "Mechanical";
		}
		if (string.Equals(normalizedValue, "전기", StringComparison.OrdinalIgnoreCase))
		{
			return "Electrical";
		}
		if (string.Equals(normalizedValue, "소방", StringComparison.OrdinalIgnoreCase))
		{
			return "FireProtection";
		}
		if (string.Equals(normalizedValue, "통합", StringComparison.OrdinalIgnoreCase))
		{
			return "Integrated";
		}
		if (string.Equals(normalizedValue, "공통", StringComparison.OrdinalIgnoreCase) || string.Equals(normalizedValue, "기타", StringComparison.OrdinalIgnoreCase))
		{
			return "Other";
		}
		string text = normalizedValue;
		switch (text)
		{
		default:
			if (Operators.CompareString(text, FamilyBrowserPolicyKey.Normalize("Architecture"), TextCompare: false) != 0)
			{
				switch (text)
				{
				default:
					if (Operators.CompareString(text, FamilyBrowserPolicyKey.Normalize("Structure"), TextCompare: false) != 0)
					{
						switch (text)
						{
						default:
							if (Operators.CompareString(text, FamilyBrowserPolicyKey.Normalize("Mechanical"), TextCompare: false) != 0)
							{
								switch (text)
								{
								default:
									if (Operators.CompareString(text, FamilyBrowserPolicyKey.Normalize("Electrical"), TextCompare: false) != 0)
									{
										switch (text)
										{
										default:
											if (Operators.CompareString(text, FamilyBrowserPolicyKey.Normalize("FireProtection"), TextCompare: false) != 0)
											{
												if (Operators.CompareString(text, "integrated", TextCompare: false) == 0 || Operators.CompareString(text, "통합", TextCompare: false) == 0 || Operators.CompareString(text, FamilyBrowserPolicyKey.Normalize("Integrated"), TextCompare: false) == 0)
												{
													return "Integrated";
												}
												switch (text)
												{
												default:
													if (Operators.CompareString(text, FamilyBrowserPolicyKey.Normalize("Other"), TextCompare: false) != 0)
													{
														return string.Empty;
													}
													goto case "other";
												case "other":
												case "common":
												case "공통":
												case "기타":
													return "Other";
												}
											}
											goto case "fire";
										case "fire":
										case "fireprotection":
										case "fire-protection":
										case "소방":
											return "FireProtection";
										}
									}
									goto case "elec";
								case "elec":
								case "electrical":
								case "전기":
									return "Electrical";
								}
							}
							goto case "mep";
						case "mep":
						case "mech":
						case "mechanical":
						case "mechanic":
						case "설비":
						case "기계":
							return "Mechanical";
						}
					}
					goto case "struct";
				case "struct":
				case "structure":
				case "structural":
				case "구조":
					return "Structure";
				}
			}
			goto case "arch";
		case "arch":
		case "architecture":
		case "architectural":
		case "건축":
			return "Architecture";
		}
	}

	private static string SaveSlotRegistration(string workspaceRoot, string slotKey, StandardLibraryRegistrationRecord registration)
	{
		RequireManagedDataRootForWrite(workspaceRoot, T("Save standard RVT registration", "표준 RVT 등록 정보 저장"));
		string safeSlotKey = FamilyBrowserPolicyKey.Normalize(slotKey);
		if (string.IsNullOrWhiteSpace(safeSlotKey))
		{
			safeSlotKey = "standard";
		}
		string registryFolder = GetRegistryFolder(workspaceRoot);
		if (string.IsNullOrWhiteSpace(registryFolder))
		{
			throw new InvalidOperationException(BuildManagedDataRootRequiredMessage(T("Save standard RVT registration", "표준 RVT 등록 정보 저장")));
		}
		string text = Path.Combine(registryFolder, "standard-library-" + safeSlotKey + ".json");
		Directory.CreateDirectory(Path.GetDirectoryName(text));
		string temporaryPath = FamilyBrowserAtomicFileService.CreateSiblingTemporaryPath(text);
		try
		{
			byte[] payload = new UTF8Encoding(false).GetBytes(PlainJsonReportWriter.Serialize(registration));
			using (FileStream stream = new FileStream(temporaryPath, FileMode.CreateNew, FileAccess.Write, FileShare.None))
			{
				stream.Write(payload, 0, payload.Length);
				stream.Flush(true);
			}
			FamilyBrowserAtomicFileService.Promote(temporaryPath, text);
		}
		finally
		{
			try
			{
				if (File.Exists(temporaryPath))
				{
					File.Delete(temporaryPath);
				}
			}
			catch
			{
			}
		}
		return text;
	}

	private static bool HasAnyPolicyRegistration(FamilyBrowserStandardPolicy policy)
	{
		if (policy == null)
		{
			return false;
		}
		if (policy.IntegratedLibrary != null && !string.IsNullOrWhiteSpace(policy.IntegratedLibrary.RegistrationPath))
		{
			return true;
		}
		if (policy.DisciplineLibraries != null && policy.DisciplineLibraries.Any([SpecialName] (FamilyBrowserStandardLibrarySlot x) => x != null && !string.IsNullOrWhiteSpace(x.RegistrationPath)))
		{
			return true;
		}
		return false;
	}

	private static FamilyBrowserStandardPolicy CreateDefaultPolicy(string currentUser)
	{
		return new FamilyBrowserStandardPolicy
		{
			Mode = "DisciplineSeparated",
			ActiveDiscipline = "Mechanical",
			IntegratedLibrary = FamilyBrowserStandardLibrarySlot.CreateIntegrated(),
			DisciplineLibraries = FamilyBrowserStandardLibrarySlot.CreateDefaultDisciplines(),
			RequestStore = FamilyBrowserRequestStoreSettings.CreateDefault(),
			PermissionExcel = FamilyBrowserPermissionExcelSettings.CreateDefault(),
			Security = CreateDefaultSecurity(currentUser),
			ProjectPolicyRules = new List<FamilyBrowserProjectPolicyRule>(),
			FileGuard = FamilyBrowserFileGuardPolicy.CreateDefault(),
			CompareDetailedSystemTypeComponents = true,
			TrackProjectElementChanges = false,
			LastUpdatedUtc = DateTime.UtcNow.ToString("O", CultureInfo.InvariantCulture),
			LastUpdatedBy = (currentUser ?? string.Empty)
		};
	}

	private static void ApplyRequestStore(FamilyBrowserStandardPolicy policy, string mode, string path, string endpoint, string updatedUtc, string currentUser)
	{
		NormalizePolicy(policy);
		policy.RequestStore.Mode = ResolveRequestStoreMode(mode);
		policy.RequestStore.Path = (path ?? string.Empty).Trim();
		policy.RequestStore.Endpoint = (endpoint ?? string.Empty).Trim();
		policy.RequestStore.LastUpdatedUtc = updatedUtc;
		policy.RequestStore.LastUpdatedBy = currentUser ?? string.Empty;
	}

	private static void ApplyPermissionExcel(FamilyBrowserStandardPolicy policy, bool enabled, string excelPath, string sheetName, string updatedUtc, string currentUser)
	{
		NormalizePolicy(policy);
		policy.PermissionExcel.Enabled = enabled;
		policy.PermissionExcel.Path = (excelPath ?? string.Empty).Trim();
		policy.PermissionExcel.SheetName = (sheetName ?? string.Empty).Trim();
		policy.PermissionExcel.LastUpdatedUtc = updatedUtc;
		policy.PermissionExcel.LastUpdatedBy = currentUser ?? string.Empty;
	}

	public static void SetRequestStore(string workspaceRoot, FamilyBrowserStandardPolicy policy, string mode, string path, string endpoint, string currentUser)
	{
		NormalizePolicy(policy);
		string mode2 = ResolveRequestStoreMode(mode);
		string path2 = (path ?? string.Empty).Trim();
		string endpoint2 = (endpoint ?? string.Empty).Trim();
		string updatedUtc = DateTime.UtcNow.ToString("O", CultureInfo.InvariantCulture);
		MutateLatestPolicy(workspaceRoot, currentUser, [SpecialName] (FamilyBrowserStandardPolicy latest) =>
		{
			ApplyRequestStore(latest, mode2, path2, endpoint2, updatedUtc, currentUser);
			ApplyRequestStore(policy, mode2, path2, endpoint2, updatedUtc, currentUser);
		});
	}

	public static void SetPermissionExcel(string workspaceRoot, FamilyBrowserStandardPolicy policy, bool enabled, string excelPath, string sheetName, string currentUser)
	{
		NormalizePolicy(policy);
		string excelPath2 = (excelPath ?? string.Empty).Trim();
		string sheetName2 = (sheetName ?? string.Empty).Trim();
		string updatedUtc = DateTime.UtcNow.ToString("O", CultureInfo.InvariantCulture);
		MutateLatestPolicy(workspaceRoot, currentUser, [SpecialName] (FamilyBrowserStandardPolicy latest) =>
		{
			ApplyPermissionExcel(latest, enabled, excelPath2, sheetName2, updatedUtc, currentUser);
			ApplyPermissionExcel(policy, enabled, excelPath2, sheetName2, updatedUtc, currentUser);
		});
	}

	public static void SetSecurityUsers(string workspaceRoot, FamilyBrowserStandardPolicy policy, string role, string rawUsers, string currentUser)
	{
		NormalizePolicy(policy);
		List<string> list = FamilyBrowserSecurityPolicyService.ParseUserList(rawUsers);
		switch (FamilyBrowserPolicyKey.Normalize(role))
		{
		case "admin":
		case "admins":
			MutateLatestPolicy(workspaceRoot, currentUser, [SpecialName] (FamilyBrowserStandardPolicy latest) =>
			{
				latest.Security.AdminUsers = list;
				latest.Security.LastUpdatedUtc = DateTime.UtcNow.ToString("O", CultureInfo.InvariantCulture);
				latest.Security.LastUpdatedBy = currentUser ?? string.Empty;
				policy.Security.AdminUsers = list;
				policy.Security.LastUpdatedUtc = latest.Security.LastUpdatedUtc;
				policy.Security.LastUpdatedBy = latest.Security.LastUpdatedBy;
			});
			break;
		case "adminprofilekeywords":
		case "admin-profile-keywords":
		case "adminkeywords":
		case "admin-keywords":
			MutateLatestPolicy(workspaceRoot, currentUser, [SpecialName] (FamilyBrowserStandardPolicy latest) =>
			{
				latest.Security.AdminProfileKeywords = list;
				latest.Security.LastUpdatedUtc = DateTime.UtcNow.ToString("O", CultureInfo.InvariantCulture);
				latest.Security.LastUpdatedBy = currentUser ?? string.Empty;
				policy.Security.AdminProfileKeywords = list;
				policy.Security.LastUpdatedUtc = latest.Security.LastUpdatedUtc;
				policy.Security.LastUpdatedBy = latest.Security.LastUpdatedBy;
			});
			break;
		case "approver":
		case "approvers":
		case "requestapprover":
		case "request-approver":
			MutateLatestPolicy(workspaceRoot, currentUser, [SpecialName] (FamilyBrowserStandardPolicy latest) =>
			{
				latest.Security.RequestApproverUsers = list;
				latest.Security.LastUpdatedUtc = DateTime.UtcNow.ToString("O", CultureInfo.InvariantCulture);
				latest.Security.LastUpdatedBy = currentUser ?? string.Empty;
				policy.Security.RequestApproverUsers = list;
				policy.Security.LastUpdatedUtc = latest.Security.LastUpdatedUtc;
				policy.Security.LastUpdatedBy = latest.Security.LastUpdatedBy;
			});
			break;
		case "readonly":
		case "read-only":
		case "viewer":
		case "viewers":
			MutateLatestPolicy(workspaceRoot, currentUser, [SpecialName] (FamilyBrowserStandardPolicy latest) =>
			{
				latest.Security.ReadOnlyUsers = list;
				latest.Security.LastUpdatedUtc = DateTime.UtcNow.ToString("O", CultureInfo.InvariantCulture);
				latest.Security.LastUpdatedBy = currentUser ?? string.Empty;
				policy.Security.ReadOnlyUsers = list;
				policy.Security.LastUpdatedUtc = latest.Security.LastUpdatedUtc;
				policy.Security.LastUpdatedBy = latest.Security.LastUpdatedBy;
			});
			break;
		default:
			throw new ArgumentException(T("Unknown security role: ", "알 수 없는 보안 역할: ") + (role ?? string.Empty), "role");
		}
	}

	public static void SetFileGuardPolicy(string workspaceRoot, FamilyBrowserStandardPolicy policy, FamilyBrowserFileGuardPolicy fileGuard, string currentUser)
	{
		if (fileGuard == null)
		{
			fileGuard = FamilyBrowserFileGuardPolicy.CreateDefault();
		}
		NormalizeFileGuard(fileGuard);
		fileGuard.LastUpdatedUtc = DateTime.UtcNow.ToString("O", CultureInfo.InvariantCulture);
		fileGuard.LastUpdatedBy = currentUser ?? string.Empty;
		FamilyBrowserFileGuardPolicy fileGuard2 = CloneFileGuardPolicy(fileGuard);
		bool fileScopedTrackingEnabled = (fileGuard2.Targets ?? new List<FamilyBrowserFileGuardTarget>())
			.Any([SpecialName] (FamilyBrowserFileGuardTarget target) => target != null && target.Enabled && target.TrackElementChanges);
		MutateLatestPolicy(workspaceRoot, currentUser, [SpecialName] (FamilyBrowserStandardPolicy latest) =>
		{
			latest.FileGuard = CloneFileGuardPolicy(fileGuard2);
			policy.FileGuard = CloneFileGuardPolicy(fileGuard2);
			latest.TrackProjectElementChanges = fileScopedTrackingEnabled;
			policy.TrackProjectElementChanges = fileScopedTrackingEnabled;
		});
	}

	public static void AddOrUpdateProjectPolicyRule(string workspaceRoot, FamilyBrowserStandardPolicy policy, FamilyBrowserProjectPolicyRule rule, string currentUser)
	{
		if (rule == null)
		{
			throw new ArgumentNullException("rule");
		}
		NormalizePolicy(policy);
		NormalizeProjectPolicyRule(rule);
		rule.LastUpdatedUtc = DateTime.UtcNow.ToString("O", CultureInfo.InvariantCulture);
		rule.LastUpdatedBy = currentUser ?? string.Empty;
		string text = FamilyBrowserPolicyKey.Normalize(rule.RuleName);
		if (string.IsNullOrWhiteSpace(text))
		{
			text = FamilyBrowserPolicyKey.Normalize(rule.MatchMode + "-" + rule.MatchValue);
		}
		FamilyBrowserProjectPolicyRule rule2 = CloneProjectPolicyRule(rule);
		MutateLatestPolicy(workspaceRoot, currentUser, [SpecialName] (FamilyBrowserStandardPolicy latest) =>
		{
			UpsertProjectPolicyRuleInPolicy(latest, CloneProjectPolicyRule(rule2), text);
			UpsertProjectPolicyRuleInPolicy(policy, CloneProjectPolicyRule(rule2), text);
		});
	}

	public static void ClearProjectPolicyRules(string workspaceRoot, FamilyBrowserStandardPolicy policy, string currentUser)
	{
		NormalizePolicy(policy);
		MutateLatestPolicy(workspaceRoot, currentUser, [SpecialName] (FamilyBrowserStandardPolicy latest) =>
		{
			latest.ProjectPolicyRules.Clear();
			policy.ProjectPolicyRules.Clear();
		});
	}

	private static void UpsertProjectPolicyRuleInPolicy(FamilyBrowserStandardPolicy policy, FamilyBrowserProjectPolicyRule rule, string key)
	{
		NormalizePolicy(policy);
		FamilyBrowserProjectPolicyRule existing = policy.ProjectPolicyRules.FirstOrDefault([SpecialName] (FamilyBrowserProjectPolicyRule x) => string.Equals(FamilyBrowserPolicyKey.Normalize(x.RuleName), key, StringComparison.OrdinalIgnoreCase));
		if (existing != null)
		{
			policy.ProjectPolicyRules.Remove(existing);
		}
		policy.ProjectPolicyRules.Add(rule);
		policy.ProjectPolicyRules = policy.ProjectPolicyRules.OrderBy<FamilyBrowserProjectPolicyRule, string>([SpecialName] (FamilyBrowserProjectPolicyRule x) => x.RuleName, StringComparer.OrdinalIgnoreCase).ToList();
	}

	private static FamilyBrowserProjectPolicyRule CloneProjectPolicyRule(FamilyBrowserProjectPolicyRule rule)
	{
		if (rule == null)
		{
			return null;
		}
		return new FamilyBrowserProjectPolicyRule
		{
			RuleName = (rule.RuleName ?? string.Empty),
			Enabled = rule.Enabled,
			MatchMode = (rule.MatchMode ?? string.Empty),
			MatchValue = (rule.MatchValue ?? string.Empty),
			PermissionPreset = (rule.PermissionPreset ?? string.Empty),
			CustomAdminUsers = new List<string>(rule.CustomAdminUsers ?? new List<string>()),
			CustomRequestApproverUsers = new List<string>(rule.CustomRequestApproverUsers ?? new List<string>()),
			CustomReadOnlyUsers = new List<string>(rule.CustomReadOnlyUsers ?? new List<string>()),
			AllowUnlistedUsersAsModelers = (rule.AllowUnlistedUsersAsModelers ?? string.Empty),
			AllowModelersToLoadFamilies = (rule.AllowModelersToLoadFamilies ?? string.Empty),
			AllowModelersToApplySystemTypes = (rule.AllowModelersToApplySystemTypes ?? string.Empty),
			AllowModelersToSubmitRequests = (rule.AllowModelersToSubmitRequests ?? string.Empty),
			LastUpdatedUtc = (rule.LastUpdatedUtc ?? string.Empty),
			LastUpdatedBy = (rule.LastUpdatedBy ?? string.Empty)
		};
	}

	private static FamilyBrowserFileGuardPolicy CloneFileGuardPolicy(FamilyBrowserFileGuardPolicy fileGuard)
	{
		if (fileGuard == null)
		{
			return FamilyBrowserFileGuardPolicy.CreateDefault();
		}
		return new FamilyBrowserFileGuardPolicy
		{
			Enabled = fileGuard.Enabled,
			RootFolder = (fileGuard.RootFolder ?? string.Empty),
			Targets = (from x in fileGuard.Targets ?? new List<FamilyBrowserFileGuardTarget>()
				where x != null
				select CloneFileGuardTarget(x)).ToList(),
			LastUpdatedUtc = (fileGuard.LastUpdatedUtc ?? string.Empty),
			LastUpdatedBy = (fileGuard.LastUpdatedBy ?? string.Empty)
		};
	}

	private static FamilyBrowserFileGuardTarget CloneFileGuardTarget(FamilyBrowserFileGuardTarget target)
	{
		if (target == null)
		{
			return null;
		}
		return new FamilyBrowserFileGuardTarget
		{
			Enabled = target.Enabled,
			FileName = (target.FileName ?? string.Empty),
			CentralPath = (target.CentralPath ?? string.Empty),
			RelativePath = (target.RelativePath ?? string.Empty),
			Discipline = (target.Discipline ?? string.Empty),
			BlockFamilyLoadAndEdit = target.BlockFamilyLoadAndEdit,
			BlockTypeChanges = target.BlockTypeChanges,
			BlockNestedOnlyStandalonePlacement = target.BlockNestedOnlyStandalonePlacement,
			TrackElementChanges = target.TrackElementChanges,
			TrackElementChangesConfigured = target.TrackElementChangesConfigured,
			LastUpdatedUtc = (target.LastUpdatedUtc ?? string.Empty),
			LastUpdatedBy = (target.LastUpdatedBy ?? string.Empty)
		};
	}

	private static void NormalizePolicy(FamilyBrowserStandardPolicy policy)
	{
		if (policy == null)
		{
			throw new ArgumentNullException("policy");
		}
		if (string.IsNullOrWhiteSpace(policy.Mode))
		{
			policy.Mode = "DisciplineSeparated";
		}
		if (string.IsNullOrWhiteSpace(policy.ActiveDiscipline))
		{
			policy.ActiveDiscipline = "Mechanical";
		}
		if (policy.IntegratedLibrary == null)
		{
			policy.IntegratedLibrary = FamilyBrowserStandardLibrarySlot.CreateIntegrated();
		}
		if (policy.DisciplineLibraries == null)
		{
			policy.DisciplineLibraries = new List<FamilyBrowserStandardLibrarySlot>();
		}
		if (policy.RequestStore == null)
		{
			policy.RequestStore = FamilyBrowserRequestStoreSettings.CreateDefault();
		}
		if (policy.PermissionExcel == null)
		{
			policy.PermissionExcel = FamilyBrowserPermissionExcelSettings.CreateDefault();
		}
		if (policy.Security == null)
		{
			policy.Security = FamilyBrowserSecurityPolicy.CreateDefault();
		}
		if (policy.ProjectPolicyRules == null)
		{
			policy.ProjectPolicyRules = new List<FamilyBrowserProjectPolicyRule>();
		}
		if (policy.FileGuard == null)
		{
			policy.FileGuard = FamilyBrowserFileGuardPolicy.CreateDefault();
		}
		if (!policy.CompareDetailedSystemTypeComponents.HasValue)
		{
			policy.CompareDetailedSystemTypeComponents = true;
		}
		if (!policy.TrackProjectElementChanges.HasValue)
		{
			policy.TrackProjectElementChanges = false;
		}
		NormalizeRequestStore(policy.RequestStore);
		NormalizePermissionExcel(policy.PermissionExcel);
		NormalizeSecurity(policy.Security);
		NormalizeProjectPolicyRules(policy.ProjectPolicyRules);
		NormalizeFileGuard(policy.FileGuard);
		EnsureDisciplineSlot(policy.DisciplineLibraries, "Architecture", "Architecture");
		EnsureDisciplineSlot(policy.DisciplineLibraries, "Structure", "Structure");
		EnsureDisciplineSlot(policy.DisciplineLibraries, "Mechanical", "Mechanical");
		EnsureDisciplineSlot(policy.DisciplineLibraries, "Electrical", "Electrical");
		EnsureDisciplineSlot(policy.DisciplineLibraries, "FireProtection", "Fire Protection");
		EnsureDisciplineSlot(policy.DisciplineLibraries, "Other", "Other");
		NormalizeSlot(policy.IntegratedLibrary, "integrated", "Integrated", "Integrated Standard RVT");
		foreach (FamilyBrowserStandardLibrarySlot slot in policy.DisciplineLibraries)
		{
			NormalizeSlot(slot, "discipline-" + FamilyBrowserPolicyKey.Normalize(slot.Discipline), slot.Discipline, slot.DisplayName);
		}
	}

	private static void EnsureDisciplineSlot(List<FamilyBrowserStandardLibrarySlot> slots, string discipline, string displayName)
	{
		if (!slots.Any([SpecialName] (FamilyBrowserStandardLibrarySlot x) => string.Equals(FamilyBrowserPolicyKey.Normalize(x.Discipline), FamilyBrowserPolicyKey.Normalize(discipline), StringComparison.OrdinalIgnoreCase)))
		{
			slots.Add(FamilyBrowserStandardLibrarySlot.CreateDiscipline(discipline, displayName));
		}
	}

	private static void NormalizeSlot(FamilyBrowserStandardLibrarySlot slot, string fallbackSlotKey, string fallbackDiscipline, string fallbackDisplayName)
	{
		if (slot != null)
		{
			if (string.IsNullOrWhiteSpace(slot.Discipline))
			{
				slot.Discipline = fallbackDiscipline;
			}
			if (string.IsNullOrWhiteSpace(slot.SlotKey))
			{
				slot.SlotKey = fallbackSlotKey;
			}
			if (string.IsNullOrWhiteSpace(slot.DisplayName))
			{
				slot.DisplayName = fallbackDisplayName;
			}
			slot.StandardListPath = (slot.StandardListPath ?? string.Empty).Trim();
			slot.StandardListSheetName = (slot.StandardListSheetName ?? string.Empty).Trim();
		}
	}

	private static void NormalizeRequestStore(FamilyBrowserRequestStoreSettings settings)
	{
		if (settings != null)
		{
			settings.Mode = ResolveRequestStoreMode(settings.Mode);
			settings.Path = (settings.Path ?? string.Empty).Trim();
			settings.Endpoint = (settings.Endpoint ?? string.Empty).Trim();
		}
	}

	private static void NormalizePermissionExcel(FamilyBrowserPermissionExcelSettings settings)
	{
		if (settings != null)
		{
			settings.Path = (settings.Path ?? string.Empty).Trim();
			settings.SheetName = (settings.SheetName ?? string.Empty).Trim();
			if (string.IsNullOrWhiteSpace(settings.Path))
			{
				settings.Enabled = false;
			}
		}
	}

	private static FamilyBrowserSecurityPolicy CreateDefaultSecurity(string currentUser)
	{
		FamilyBrowserSecurityPolicy security = FamilyBrowserSecurityPolicy.CreateDefault();
		string user = (currentUser ?? string.Empty).Trim();
		string activeUser = FamilyBrowserSecurityPolicyService.ResolveCurrentUserIdentity();
		if (!string.IsNullOrWhiteSpace(activeUser))
		{
			user = activeUser;
		}
		if (!string.IsNullOrWhiteSpace(user))
		{
			security.AdminUsers.Add(user);
		}
		security.LastUpdatedUtc = DateTime.UtcNow.ToString("O", CultureInfo.InvariantCulture);
		security.LastUpdatedBy = user;
		return security;
	}

	private static void NormalizeSecurity(FamilyBrowserSecurityPolicy security)
	{
		if (security != null)
		{
			if (security.AdminUsers == null)
			{
				security.AdminUsers = new List<string>();
			}
			if (security.AdminProfileKeywords == null)
			{
				security.AdminProfileKeywords = new List<string>();
			}
			if (security.RequestApproverUsers == null)
			{
				security.RequestApproverUsers = new List<string>();
			}
			if (security.ReadOnlyUsers == null)
			{
				security.ReadOnlyUsers = new List<string>();
			}
			security.AdminUsers = FamilyBrowserSecurityPolicyService.ParseUserList(FamilyBrowserSecurityPolicyService.FormatUserList(security.AdminUsers));
			security.AdminProfileKeywords = FamilyBrowserSecurityPolicyService.ParseUserList(FamilyBrowserSecurityPolicyService.FormatUserList(security.AdminProfileKeywords));
			security.RequestApproverUsers = FamilyBrowserSecurityPolicyService.ParseUserList(FamilyBrowserSecurityPolicyService.FormatUserList(security.RequestApproverUsers));
			security.ReadOnlyUsers = FamilyBrowserSecurityPolicyService.ParseUserList(FamilyBrowserSecurityPolicyService.FormatUserList(security.ReadOnlyUsers));
		}
	}

	private static void NormalizeProjectPolicyRules(List<FamilyBrowserProjectPolicyRule> rules)
	{
		if (rules == null)
		{
			return;
		}
		foreach (FamilyBrowserProjectPolicyRule rule in rules)
		{
			NormalizeProjectPolicyRule(rule);
		}
	}

	private static void NormalizeProjectPolicyRule(FamilyBrowserProjectPolicyRule rule)
	{
		if (rule != null)
		{
			rule.RuleName = (rule.RuleName ?? string.Empty).Trim();
			rule.MatchValue = (rule.MatchValue ?? string.Empty).Trim();
			rule.MatchMode = ResolveProjectPolicyMatchMode(rule.MatchMode);
			rule.PermissionPreset = ResolveProjectPolicyPreset(rule.PermissionPreset);
			rule.AllowUnlistedUsersAsModelers = ResolveProjectPolicyOverride(rule.AllowUnlistedUsersAsModelers);
			rule.AllowModelersToLoadFamilies = ResolveProjectPolicyOverride(rule.AllowModelersToLoadFamilies);
			rule.AllowModelersToApplySystemTypes = ResolveProjectPolicyOverride(rule.AllowModelersToApplySystemTypes);
			rule.AllowModelersToSubmitRequests = ResolveProjectPolicyOverride(rule.AllowModelersToSubmitRequests);
			if (rule.CustomAdminUsers == null)
			{
				rule.CustomAdminUsers = new List<string>();
			}
			if (rule.CustomRequestApproverUsers == null)
			{
				rule.CustomRequestApproverUsers = new List<string>();
			}
			if (rule.CustomReadOnlyUsers == null)
			{
				rule.CustomReadOnlyUsers = new List<string>();
			}
			rule.CustomAdminUsers = FamilyBrowserSecurityPolicyService.ParseUserList(FamilyBrowserSecurityPolicyService.FormatUserList(rule.CustomAdminUsers));
			rule.CustomRequestApproverUsers = FamilyBrowserSecurityPolicyService.ParseUserList(FamilyBrowserSecurityPolicyService.FormatUserList(rule.CustomRequestApproverUsers));
			rule.CustomReadOnlyUsers = FamilyBrowserSecurityPolicyService.ParseUserList(FamilyBrowserSecurityPolicyService.FormatUserList(rule.CustomReadOnlyUsers));
		}
	}

	private static void NormalizeFileGuard(FamilyBrowserFileGuardPolicy fileGuard)
	{
		if (fileGuard == null)
		{
			return;
		}
		fileGuard.RootFolder = (fileGuard.RootFolder ?? string.Empty).Trim();
		if (fileGuard.Targets == null)
		{
			fileGuard.Targets = new List<FamilyBrowserFileGuardTarget>();
		}
		foreach (FamilyBrowserFileGuardTarget target in fileGuard.Targets)
		{
			NormalizeFileGuardTarget(target);
		}
		fileGuard.Targets = fileGuard.Targets.Where([SpecialName] (FamilyBrowserFileGuardTarget x) => x != null && !string.IsNullOrWhiteSpace(x.FileName)).OrderBy<FamilyBrowserFileGuardTarget, string>([SpecialName] (FamilyBrowserFileGuardTarget x) => x.FileName, StringComparer.OrdinalIgnoreCase).ThenBy<FamilyBrowserFileGuardTarget, string>([SpecialName] (FamilyBrowserFileGuardTarget x) => x.RelativePath ?? string.Empty, StringComparer.OrdinalIgnoreCase)
			.ToList();
	}

	private static void NormalizeFileGuardTarget(FamilyBrowserFileGuardTarget target)
	{
		if (target != null)
		{
			target.FileName = (target.FileName ?? string.Empty).Trim();
			target.CentralPath = (target.CentralPath ?? string.Empty).Trim();
			target.RelativePath = (target.RelativePath ?? string.Empty).Trim();
			target.Discipline = (target.Discipline ?? string.Empty).Trim();
			if (!target.TrackElementChangesConfigured)
			{
				target.TrackElementChanges = true;
				target.TrackElementChangesConfigured = true;
			}
		}
	}

	private static string ResolveProjectPolicyMatchMode(string matchMode)
	{
		switch (FamilyBrowserPolicyKey.Normalize(matchMode))
		{
		case "any":
			return "Any";
		case "exactcentralpath":
		case "exact-central-path":
			return "ExactCentralPath";
		case "exactmodelpath":
		case "exact-model-path":
			return "ExactModelPath";
		case "modelpathcontains":
		case "model-path-contains":
			return "ModelPathContains";
		case "projecttitlecontains":
		case "project-title-contains":
		case "titlecontains":
		case "title-contains":
			return "ProjectTitleContains";
		default:
			return "CentralPathContains";
		}
	}

	private static string ResolveProjectPolicyPreset(string preset)
	{
		switch (FamilyBrowserPolicyKey.Normalize(preset))
		{
		case "standardmodeler":
		case "standard-modeler":
		case "modelerload":
		case "modeler-load":
			return "StandardModeler";
		case "requestonly":
		case "request-only":
		case "request":
			return "RequestOnly";
		case "readonly":
		case "read-only":
		case "locked":
			return "ReadOnly";
		default:
			return "Inherit";
		}
	}

	private static string ResolveProjectPolicyOverride(string value)
	{
		switch (FamilyBrowserPolicyKey.Normalize(value))
		{
		case "allow":
		case "true":
		case "yes":
		case "y":
			return "Allow";
		case "deny":
		case "false":
		case "no":
		case "n":
			return "Deny";
		default:
			return "Inherit";
		}
	}

	private static string ResolveRequestStoreMode(string mode)
	{
		switch (FamilyBrowserPolicyKey.Normalize(mode))
		{
		case "networkshare":
		case "network-share":
		case "network":
		case "unc":
		case "share":
			return "NetworkShare";
		case "sharepoint":
		case "share-point":
		case "m365":
		case "office365":
			return "SharePoint";
		case "cloudstorage":
		case "cloud-storage":
		case "cloud":
			return "CloudStorage";
		case "api":
		case "server":
		case "database":
		case "db":
			return "Api";
		default:
			return "Local";
		}
	}

	private static string ResolveUsableManagedPolicyPath()
	{
		string managedPolicyPath = FamilyBrowserMachineConfigStore.ResolveManagedPolicyPath();
		if (IsManagedPolicyPathUsableFast(managedPolicyPath))
		{
			return managedPolicyPath;
		}
		return string.Empty;
	}

	private static bool IsManagedPolicyPathUsableFast(string managedPolicyPath)
	{
		if (string.IsNullOrWhiteSpace(managedPolicyPath))
		{
			return false;
		}
		string normalizedPath = managedPolicyPath.Trim();
		DateTime now = DateTime.UtcNow;
		object managedPolicyProbeSyncRoot = ManagedPolicyProbeSyncRoot;
		ObjectFlowControl.CheckForSyncLockOnValueType(managedPolicyProbeSyncRoot);
		bool lockTaken = false;
		try
		{
			Monitor.Enter(managedPolicyProbeSyncRoot, ref lockTaken);
			if (string.Equals(_lastManagedPolicyProbePath, normalizedPath, StringComparison.OrdinalIgnoreCase) && (now - _lastManagedPolicyProbeUtc).TotalSeconds < 30.0)
			{
				return !string.IsNullOrWhiteSpace(_lastManagedPolicyProbeResult);
			}
		}
		finally
		{
			if (lockTaken)
			{
				Monitor.Exit(managedPolicyProbeSyncRoot);
			}
		}
		bool usable = ProbeManagedPolicyPath(normalizedPath);
		object managedPolicyProbeSyncRoot2 = ManagedPolicyProbeSyncRoot;
		ObjectFlowControl.CheckForSyncLockOnValueType(managedPolicyProbeSyncRoot2);
		bool lockTaken2 = false;
		try
		{
			Monitor.Enter(managedPolicyProbeSyncRoot2, ref lockTaken2);
			_lastManagedPolicyProbePath = normalizedPath;
			_lastManagedPolicyProbeUtc = now;
			_lastManagedPolicyProbeResult = (usable ? normalizedPath : string.Empty);
		}
		finally
		{
			if (lockTaken2)
			{
				Monitor.Exit(managedPolicyProbeSyncRoot2);
			}
		}
		return usable;
	}

	private static bool ProbeManagedPolicyPath(string managedPolicyPath)
	{
		bool ProbeManagedPolicyPath;
		try
		{
			Task<bool> probe = Task.Factory.StartNew([SpecialName] () =>
			{
				bool result;
				try
				{
					if (File.Exists(managedPolicyPath))
					{
						result = true;
					}
					else
					{
						string directoryName = Path.GetDirectoryName(managedPolicyPath);
						result = !string.IsNullOrWhiteSpace(directoryName) && Directory.Exists(directoryName);
					}
				}
				catch (Exception projectError2)
				{
					ProjectData.SetProjectError(projectError2);
					result = false;
					ProjectData.ClearProjectError();
				}
				return result;
			}, CancellationToken.None, TaskCreationOptions.DenyChildAttach, TaskScheduler.Default);
			ProbeManagedPolicyPath = probe.Wait(800) && probe.Result;
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ProbeManagedPolicyPath = false;
			ProjectData.ClearProjectError();
		}
		return ProbeManagedPolicyPath;
	}

	private static void EnsureParentFolder(string filePath)
	{
		string? directoryName = Path.GetDirectoryName(filePath);
		if (string.IsNullOrWhiteSpace(directoryName))
		{
			throw new InvalidOperationException(T("Policy file path must include a folder.", "정책 파일 경로에는 폴더가 포함되어야 합니다."));
		}
		Directory.CreateDirectory(directoryName);
	}

	private static string ResolveManagedPolicyFolder(string managedPolicyPath, string workspaceRoot)
	{
		string folder = Path.GetDirectoryName(managedPolicyPath);
		if (string.IsNullOrWhiteSpace(folder))
		{
			return string.Empty;
		}
		return folder;
	}

	private static string ResolveManagedDataRoot(string policyFolder)
	{
		if (string.IsNullOrWhiteSpace(policyFolder))
		{
			return string.Empty;
		}
		string trimmed = policyFolder.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
		if (string.Equals(Path.GetFileName(trimmed), "Config", StringComparison.OrdinalIgnoreCase))
		{
			DirectoryInfo parent = Directory.GetParent(trimmed);
			if (parent != null && !string.IsNullOrWhiteSpace(parent.FullName))
			{
				return parent.FullName;
			}
		}
		return policyFolder;
	}

	private static void MigrateLegacyManagedRegistryFolder(string workspaceRoot, string managedPolicyPath, string registryFolder)
	{
		if (string.IsNullOrWhiteSpace(registryFolder))
		{
			return;
		}
		string legacyFolder = Path.Combine(ResolveManagedPolicyFolder(managedPolicyPath, workspaceRoot), "Registry");
		if (string.IsNullOrWhiteSpace(legacyFolder) || string.Equals(legacyFolder, registryFolder, StringComparison.OrdinalIgnoreCase) || !Directory.Exists(legacyFolder))
		{
			return;
		}
		Directory.CreateDirectory(registryFolder);
		string[] files = Directory.GetFiles(legacyFolder, "*.json", SearchOption.TopDirectoryOnly);
		foreach (string sourcePath in files)
		{
			string targetPath = Path.Combine(registryFolder, Path.GetFileName(sourcePath));
			if (!File.Exists(targetPath))
			{
				File.Copy(sourcePath, targetPath, overwrite: false);
			}
		}
	}

	private static string GetLocalRegistryFolder(string workspaceRoot)
	{
		return GetUnavailableManagedDataFolder(workspaceRoot, "Registry");
	}

	private static string GetUnavailableManagedDataFolder(string workspaceRoot, string folderName)
	{
		string root = (string.IsNullOrWhiteSpace(workspaceRoot) ? HostWorkspacePathResolver.ResolveRoot() : workspaceRoot);
		if (string.IsNullOrWhiteSpace(root))
		{
			return string.Empty;
		}
		string dataRoot = Path.Combine(root, "NoManagedDataRoot");
		if (string.IsNullOrWhiteSpace(folderName))
		{
			return dataRoot;
		}
		return Path.Combine(dataRoot, folderName);
	}
}
