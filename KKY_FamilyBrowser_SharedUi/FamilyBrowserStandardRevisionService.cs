using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.Serialization.Json;
using System.Security.Cryptography;
using System.Text;
using Microsoft.Win32.SafeHandles;

public sealed class FamilyBrowserStandardRevisionManifest
{
    public int SchemaVersion { get; set; }
    public string SourceId { get; set; }
    public string StandardRvtPath { get; set; }
    public string CanonicalPath { get; set; }
    public string FileIdentity { get; set; }
    public string BaselineAtUtc { get; set; }
    public string SourceFileLastWriteUtc { get; set; }
    public long SourceFileLength { get; set; }
    public string RevisionHash { get; set; }
    public string HashMode { get; set; }
    public string SnapshotPath { get; set; }
    public string SnapshotAtUtc { get; set; }
    public string RecordedBy { get; set; }

    public FamilyBrowserStandardRevisionManifest()
    {
        SchemaVersion = 1;
        SourceId = string.Empty;
        StandardRvtPath = string.Empty;
        CanonicalPath = string.Empty;
        FileIdentity = string.Empty;
        BaselineAtUtc = string.Empty;
        SourceFileLastWriteUtc = string.Empty;
        RevisionHash = string.Empty;
        HashMode = string.Empty;
        SnapshotPath = string.Empty;
        SnapshotAtUtc = string.Empty;
        RecordedBy = string.Empty;
    }
}

public sealed class FamilyBrowserStandardRevisionState
{
    public string SourceId { get; set; }
    public string StandardRvtPath { get; set; }
    public string CanonicalPath { get; set; }
    public string FileIdentity { get; set; }
    public string StateCode { get; set; }
    public string CheckedAtUtc { get; set; }
    public string BaselineAtUtc { get; set; }
    public string RecordedLastWriteUtc { get; set; }
    public long RecordedLength { get; set; }
    public string CurrentLastWriteUtc { get; set; }
    public long CurrentLength { get; set; }
    public string RecordedRevisionHash { get; set; }
    public string CurrentRevisionHash { get; set; }
    public string HashMode { get; set; }
    public string SnapshotPath { get; set; }
    public string SnapshotAtUtc { get; set; }
    public bool Changed { get; set; }
    public bool Unavailable { get; set; }
    public bool BaselineMissing { get; set; }
    public bool PathAliasMatched { get; set; }
    public string Reason { get; set; }
    public string ErrorMessage { get; set; }

    public bool BlocksStandardUse
    {
        get { return Changed || Unavailable || BaselineMissing || !string.IsNullOrWhiteSpace(ErrorMessage); }
    }

    public FamilyBrowserStandardRevisionState()
    {
        SourceId = string.Empty;
        StandardRvtPath = string.Empty;
        CanonicalPath = string.Empty;
        FileIdentity = string.Empty;
        StateCode = "NotChecked";
        CheckedAtUtc = string.Empty;
        BaselineAtUtc = string.Empty;
        RecordedLastWriteUtc = string.Empty;
        CurrentLastWriteUtc = string.Empty;
        RecordedRevisionHash = string.Empty;
        CurrentRevisionHash = string.Empty;
        HashMode = string.Empty;
        SnapshotPath = string.Empty;
        SnapshotAtUtc = string.Empty;
        Reason = string.Empty;
        ErrorMessage = string.Empty;
    }
}

public static class FamilyBrowserPathIdentityService
{
    private const uint FileReadAttributes = 0x80;
    private const uint FileShareRead = 0x1;
    private const uint FileShareWrite = 0x2;
    private const uint FileShareDelete = 0x4;
    private const uint OpenExisting = 3;
    private const uint FileFlagBackupSemantics = 0x02000000;
    private const int UniversalNameInfoLevel = 1;
    private const int ErrorMoreData = 234;
    private const int CanonicalPathCacheLimit = 2048;
    private static readonly object CanonicalPathCacheSyncRoot = new object();
    private static readonly Dictionary<string, string> CanonicalPathCache = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

    [StructLayout(LayoutKind.Sequential)]
    private struct ByHandleFileInformation
    {
        public uint FileAttributes;
        public System.Runtime.InteropServices.ComTypes.FILETIME CreationTime;
        public System.Runtime.InteropServices.ComTypes.FILETIME LastAccessTime;
        public System.Runtime.InteropServices.ComTypes.FILETIME LastWriteTime;
        public uint VolumeSerialNumber;
        public uint FileSizeHigh;
        public uint FileSizeLow;
        public uint NumberOfLinks;
        public uint FileIndexHigh;
        public uint FileIndexLow;
    }

    [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    private static extern SafeFileHandle CreateFile(string fileName, uint desiredAccess, uint shareMode, IntPtr securityAttributes, uint creationDisposition, uint flagsAndAttributes, IntPtr templateFile);

    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern bool GetFileInformationByHandle(SafeFileHandle file, out ByHandleFileInformation information);

    [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    private static extern uint GetFinalPathNameByHandle(SafeFileHandle file, StringBuilder path, uint pathLength, uint flags);

    [DllImport("mpr.dll", CharSet = CharSet.Unicode)]
    private static extern int WNetGetUniversalName(string localPath, int infoLevel, IntPtr buffer, ref int bufferSize);

    public static string GetComparableIdentity(string path)
    {
        if (string.IsNullOrWhiteSpace(path))
        {
            return string.Empty;
        }
        string identity = GetFileIdentity(path);
        if (!string.IsNullOrWhiteSpace(identity))
        {
            return "FILE:" + identity;
        }
        return "PATH:" + NormalizePath(GetUniversalPath(path)).ToUpperInvariant();
    }

    public static string GetStablePathIdentity(string path)
    {
        if (string.IsNullOrWhiteSpace(path))
        {
            return string.Empty;
        }
        string canonical = GetCanonicalPath(path);
        if (string.IsNullOrWhiteSpace(canonical))
        {
            canonical = NormalizePath(GetUniversalPath(path));
        }
        return string.IsNullOrWhiteSpace(canonical) ? string.Empty : "PATH:" + canonical.ToUpperInvariant();
    }

    public static string GetFileIdentity(string path)
    {
        SafeFileHandle handle = null;
        try
        {
            string normalized = NormalizePath(path);
            if (string.IsNullOrWhiteSpace(normalized) || !File.Exists(normalized))
            {
                return string.Empty;
            }
            handle = CreateFile(normalized, FileReadAttributes, FileShareRead | FileShareWrite | FileShareDelete, IntPtr.Zero, OpenExisting, FileFlagBackupSemantics, IntPtr.Zero);
            if (handle == null || handle.IsInvalid)
            {
                return string.Empty;
            }
            ByHandleFileInformation info;
            if (!GetFileInformationByHandle(handle, out info))
            {
                return string.Empty;
            }
            ulong index = ((ulong)info.FileIndexHigh << 32) | info.FileIndexLow;
            if (info.VolumeSerialNumber == 0 && index == 0)
            {
                return string.Empty;
            }
            return info.VolumeSerialNumber.ToString("X8", CultureInfo.InvariantCulture) + ":" + index.ToString("X16", CultureInfo.InvariantCulture);
        }
        catch
        {
            return string.Empty;
        }
        finally
        {
            if (handle != null)
            {
                handle.Dispose();
            }
        }
    }

    public static string GetCanonicalPath(string path)
    {
        SafeFileHandle handle = null;
        try
        {
            string normalized = NormalizePath(path);
            if (string.IsNullOrWhiteSpace(normalized))
            {
                return string.Empty;
            }
            string cached;
            lock (CanonicalPathCacheSyncRoot)
            {
                if (CanonicalPathCache.TryGetValue(normalized, out cached))
                {
                    return cached;
                }
            }
            if (!File.Exists(normalized) && !Directory.Exists(normalized))
            {
                return NormalizePath(GetUniversalPath(path));
            }
            handle = CreateFile(normalized, FileReadAttributes, FileShareRead | FileShareWrite | FileShareDelete, IntPtr.Zero, OpenExisting, FileFlagBackupSemantics, IntPtr.Zero);
            if (handle == null || handle.IsInvalid)
            {
                return NormalizePath(GetUniversalPath(path));
            }
            StringBuilder builder = new StringBuilder(4096);
            uint length = GetFinalPathNameByHandle(handle, builder, (uint)builder.Capacity, 0);
            if (length == 0 || length >= builder.Capacity)
            {
                return NormalizePath(GetUniversalPath(path));
            }
            string finalPath = builder.ToString();
            if (finalPath.StartsWith("\\\\?\\UNC\\", StringComparison.OrdinalIgnoreCase))
            {
                finalPath = "\\\\" + finalPath.Substring(8);
            }
            else if (finalPath.StartsWith("\\\\?\\", StringComparison.OrdinalIgnoreCase))
            {
                finalPath = finalPath.Substring(4);
            }
            string canonical = NormalizePath(finalPath);
            if (!string.IsNullOrWhiteSpace(canonical))
            {
                lock (CanonicalPathCacheSyncRoot)
                {
                    if (CanonicalPathCache.Count >= CanonicalPathCacheLimit)
                    {
                        CanonicalPathCache.Clear();
                    }
                    CanonicalPathCache[normalized] = canonical;
                }
            }
            return canonical;
        }
        catch
        {
            return NormalizePath(GetUniversalPath(path));
        }
        finally
        {
            if (handle != null)
            {
                handle.Dispose();
            }
        }
    }

    public static string NormalizePath(string value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return string.Empty;
        }
        string expanded = Environment.ExpandEnvironmentVariables(value.Trim());
        if (LooksLikeNonFileModelPath(expanded))
        {
            return expanded.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
        }
        try
        {
            return Path.GetFullPath(expanded).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
        }
        catch
        {
            return expanded.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
        }
    }

    private static bool LooksLikeNonFileModelPath(string value)
    {
        int schemeSeparator = (value ?? string.Empty).IndexOf("://", StringComparison.Ordinal);
        return schemeSeparator > 0 && !value.StartsWith("file://", StringComparison.OrdinalIgnoreCase);
    }

    private static string GetUniversalPath(string path)
    {
        string normalized = NormalizePath(path);
        if (string.IsNullOrWhiteSpace(normalized) || normalized.StartsWith("\\\\", StringComparison.Ordinal))
        {
            return normalized;
        }
        IntPtr buffer = IntPtr.Zero;
        try
        {
            int size = 0;
            int first = WNetGetUniversalName(normalized, UniversalNameInfoLevel, IntPtr.Zero, ref size);
            if (first != ErrorMoreData || size <= IntPtr.Size)
            {
                return normalized;
            }
            buffer = Marshal.AllocHGlobal(size);
            int result = WNetGetUniversalName(normalized, UniversalNameInfoLevel, buffer, ref size);
            if (result != 0)
            {
                return normalized;
            }
            IntPtr stringPointer = Marshal.ReadIntPtr(buffer);
            return Marshal.PtrToStringUni(stringPointer) ?? normalized;
        }
        catch
        {
            return normalized;
        }
        finally
        {
            if (buffer != IntPtr.Zero)
            {
                Marshal.FreeHGlobal(buffer);
            }
        }
    }
}

public static class FamilyBrowserStandardRevisionService
{
    private const int SchemaVersion = 1;
    private const int SampleSize = 1024 * 1024;
    private const string HashMode = "SHA256-SAMPLED-V1";
    private static readonly object SyncRoot = new object();

    public static string BuildCurrentRevisionToken(FamilyBrowserStandardRevisionState state)
    {
        if (state == null)
        {
            return string.Empty;
        }
        string identity = !string.IsNullOrWhiteSpace(state.FileIdentity)
            ? "FILE:" + state.FileIdentity.Trim().ToUpperInvariant()
            : "PATH:" + FamilyBrowserPathIdentityService.NormalizePath(state.CanonicalPath ?? state.StandardRvtPath).ToUpperInvariant();
        return string.Join("|", new string[]
        {
            state.SourceId ?? string.Empty,
            identity,
            state.CurrentLastWriteUtc ?? string.Empty,
            state.CurrentLength.ToString(CultureInfo.InvariantCulture),
            state.CurrentRevisionHash ?? string.Empty,
            FamilyBrowserPathIdentityService.NormalizePath(state.SnapshotPath ?? string.Empty),
            state.SnapshotAtUtc ?? string.Empty
        });
    }

    public static bool IsSameCurrentRevision(FamilyBrowserStandardRevisionState before, FamilyBrowserStandardRevisionState after)
    {
        if (before == null || after == null || before.BlocksStandardUse || after.BlocksStandardUse)
        {
            return false;
        }
        string beforeToken = BuildCurrentRevisionToken(before);
        string afterToken = BuildCurrentRevisionToken(after);
        return !string.IsNullOrWhiteSpace(beforeToken)
            && string.Equals(beforeToken, afterToken, StringComparison.OrdinalIgnoreCase);
    }

    public static bool MatchesSnapshotGeneration(FamilyBrowserStandardRevisionState state, string snapshotPath, string snapshotAtUtc)
    {
        if (state == null || state.BlocksStandardUse || string.IsNullOrWhiteSpace(state.SnapshotPath) || string.IsNullOrWhiteSpace(snapshotPath))
        {
            return false;
        }
        if (!PathsEqual(state.SnapshotPath, snapshotPath))
        {
            return false;
        }
        DateTime recordedAtUtc;
        DateTime expectedAtUtc;
        if (TryParseUtc(state.SnapshotAtUtc, out recordedAtUtc) && TryParseUtc(snapshotAtUtc, out expectedAtUtc))
        {
            return recordedAtUtc == expectedAtUtc;
        }
        return !string.IsNullOrWhiteSpace(state.SnapshotAtUtc)
            && string.Equals(state.SnapshotAtUtc.Trim(), (snapshotAtUtc ?? string.Empty).Trim(), StringComparison.OrdinalIgnoreCase);
    }

    public static FamilyBrowserStandardRevisionState Probe(string workspaceRoot, StandardLibraryRegistrationRecord registration, bool computeRevisionHash)
    {
        FamilyBrowserStandardRevisionState state = CreateState(registration);
        if (registration == null)
        {
            state.StateCode = "NotRegistered";
            state.BaselineMissing = true;
            state.Reason = "Standard RVT registration is missing.";
            return state;
        }
        string path = ResolveSourcePath(registration);
        state.StandardRvtPath = path;
        try
        {
            if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
            {
                state.StateCode = "Unavailable";
                state.Unavailable = true;
                state.Reason = "The registered Standard RVT source cannot be found.";
                return state;
            }
            FileInfo file = new FileInfo(path);
            state.CanonicalPath = FamilyBrowserPathIdentityService.GetCanonicalPath(path);
            state.FileIdentity = FamilyBrowserPathIdentityService.GetFileIdentity(path);
            state.CurrentLastWriteUtc = file.LastWriteTimeUtc.ToString("O", CultureInfo.InvariantCulture);
            state.CurrentLength = file.Length;

            FamilyBrowserStandardRevisionManifest manifest = LoadManifest(workspaceRoot, registration.SourceId);
            if (manifest == null)
            {
                DateTime registeredStamp;
                bool stampMatches = TryParseUtc(registration.SourceFileLastWriteUtc, out registeredStamp) && Math.Abs((file.LastWriteTimeUtc - registeredStamp).TotalSeconds) <= 1.0;
                bool lengthMatches = registration.SourceFileLength <= 0L || registration.SourceFileLength == file.Length;
                if (!stampMatches || !lengthMatches)
                {
                    state.StateCode = "Changed";
                    state.Changed = true;
                    state.BaselineMissing = true;
                    state.RecordedLastWriteUtc = registration.SourceFileLastWriteUtc ?? string.Empty;
                    state.RecordedLength = registration.SourceFileLength;
                    state.Reason = "The Standard RVT changed after the registered scan; a revision manifest has not been established yet.";
                    return state;
                }
                manifest = BuildManifest(registration, file, computeRevisionHash ? ComputeRevisionHash(path) : string.Empty, Environment.UserName);
                SaveManifest(workspaceRoot, manifest);
            }

            ApplyManifestToState(state, manifest);
            state.PathAliasMatched = !PathsEqual(manifest.StandardRvtPath, path) && !string.IsNullOrWhiteSpace(manifest.FileIdentity) && string.Equals(manifest.FileIdentity, state.FileIdentity, StringComparison.OrdinalIgnoreCase);
            bool identityChanged = !string.IsNullOrWhiteSpace(manifest.FileIdentity) && !string.IsNullOrWhiteSpace(state.FileIdentity) && !string.Equals(manifest.FileIdentity, state.FileIdentity, StringComparison.OrdinalIgnoreCase);
            DateTime baselineStamp;
            bool timeChanged = TryParseUtc(manifest.SourceFileLastWriteUtc, out baselineStamp) && Math.Abs((file.LastWriteTimeUtc - baselineStamp).TotalSeconds) > 1.0;
            bool lengthChanged = manifest.SourceFileLength > 0L && manifest.SourceFileLength != file.Length;
            if (identityChanged || timeChanged || lengthChanged)
            {
                state.StateCode = "Changed";
                state.Changed = true;
                state.Reason = identityChanged ? "The Standard RVT file identity changed after the last scan." : "The Standard RVT modified time or file size changed after the last scan.";
                return state;
            }
            if (computeRevisionHash)
            {
                state.CurrentRevisionHash = ComputeRevisionHash(path);
                if (!string.IsNullOrWhiteSpace(manifest.RevisionHash) && !string.Equals(manifest.RevisionHash, state.CurrentRevisionHash, StringComparison.OrdinalIgnoreCase))
                {
                    state.StateCode = "Changed";
                    state.Changed = true;
                    state.Reason = "The Standard RVT content revision changed even though its file stamp was unchanged.";
                    return state;
                }
                if (string.IsNullOrWhiteSpace(manifest.RevisionHash) && !string.IsNullOrWhiteSpace(state.CurrentRevisionHash))
                {
                    manifest.RevisionHash = state.CurrentRevisionHash;
                    manifest.HashMode = HashMode;
                    SaveManifest(workspaceRoot, manifest);
                    ApplyManifestToState(state, manifest);
                }
            }
            state.StateCode = "Current";
            state.Reason = state.PathAliasMatched ? "The registered path alias resolves to the same Standard RVT file." : "The Standard RVT matches the last registered scan.";
            return state;
        }
        catch (Exception ex)
        {
            state.StateCode = "Error";
            state.ErrorMessage = ex.Message;
            state.Reason = "The Standard RVT revision could not be verified.";
            return state;
        }
    }

    public static FamilyBrowserStandardRevisionState RecordBaseline(string workspaceRoot, StandardLibraryRegistrationRecord registration, string recordedBy)
    {
        FamilyBrowserStandardRevisionState state = CreateState(registration);
        if (registration == null)
        {
            state.StateCode = "NotRegistered";
            state.BaselineMissing = true;
            return state;
        }
        string path = ResolveSourcePath(registration);
        try
        {
            if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
            {
                state.StateCode = "Unavailable";
                state.Unavailable = true;
                state.Reason = "The source was unavailable while recording the scan revision.";
                return state;
            }
            FileInfo file = new FileInfo(path);
            DateTime registrationStamp;
            bool stampMatches = !TryParseUtc(registration.SourceFileLastWriteUtc, out registrationStamp) || Math.Abs((file.LastWriteTimeUtc - registrationStamp).TotalSeconds) <= 1.0;
            bool lengthMatches = registration.SourceFileLength <= 0L || registration.SourceFileLength == file.Length;
            if (!stampMatches || !lengthMatches)
            {
                state.StateCode = "Changed";
                state.Changed = true;
                state.Reason = "The source changed before the scan registration could be committed, so the new revision was not trusted.";
                return state;
            }
            FamilyBrowserStandardRevisionManifest manifest = BuildManifest(registration, file, ComputeRevisionHash(path), recordedBy);
            SaveManifest(workspaceRoot, manifest);
            return Probe(workspaceRoot, registration, true);
        }
        catch (Exception ex)
        {
            state.StateCode = "Error";
            state.ErrorMessage = ex.Message;
            state.Reason = "The scan completed, but its Standard RVT revision manifest could not be recorded.";
            return state;
        }
    }

    public static string GetManifestPath(string workspaceRoot, string sourceId)
    {
        string safeSource = SafeFileName(string.IsNullOrWhiteSpace(sourceId) ? "standard" : sourceId);
        return Path.Combine(FamilyBrowserStandardPolicyStore.GetDataFolder(workspaceRoot, "StandardRevisionManifests"), "standard-rvt-revision-" + safeSource + ".json");
    }

    public static string ComputeRevisionHash(string path)
    {
        if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
        {
            return string.Empty;
        }
        using (FileStream stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete))
        using (SHA256 sha = SHA256.Create())
        using (MemoryStream samples = new MemoryStream())
        {
            byte[] lengthBytes = BitConverter.GetBytes(stream.Length);
            samples.Write(lengthBytes, 0, lengthBytes.Length);
            if (stream.Length <= SampleSize * 3L)
            {
                CopyRange(stream, samples, 0L, stream.Length);
            }
            else
            {
                CopyRange(stream, samples, 0L, SampleSize);
                CopyRange(stream, samples, Math.Max(0L, stream.Length / 2L - SampleSize / 2L), SampleSize);
                CopyRange(stream, samples, Math.Max(0L, stream.Length - SampleSize), SampleSize);
            }
            byte[] hash = sha.ComputeHash(samples.ToArray());
            return string.Concat(hash.Select(delegate(byte x) { return x.ToString("x2", CultureInfo.InvariantCulture); }));
        }
    }

    private static FamilyBrowserStandardRevisionState CreateState(StandardLibraryRegistrationRecord registration)
    {
        return new FamilyBrowserStandardRevisionState
        {
            SourceId = registration == null ? string.Empty : registration.SourceId ?? string.Empty,
            StandardRvtPath = registration == null ? string.Empty : ResolveSourcePath(registration),
            CheckedAtUtc = DateTime.UtcNow.ToString("O", CultureInfo.InvariantCulture)
        };
    }

    private static FamilyBrowserStandardRevisionManifest BuildManifest(StandardLibraryRegistrationRecord registration, FileInfo file, string revisionHash, string recordedBy)
    {
        string path = ResolveSourcePath(registration);
        return new FamilyBrowserStandardRevisionManifest
        {
            SchemaVersion = SchemaVersion,
            SourceId = registration.SourceId ?? string.Empty,
            StandardRvtPath = path,
            CanonicalPath = FamilyBrowserPathIdentityService.GetCanonicalPath(path),
            FileIdentity = FamilyBrowserPathIdentityService.GetFileIdentity(path),
            BaselineAtUtc = DateTime.UtcNow.ToString("O", CultureInfo.InvariantCulture),
            SourceFileLastWriteUtc = file.LastWriteTimeUtc.ToString("O", CultureInfo.InvariantCulture),
            SourceFileLength = file.Length,
            RevisionHash = revisionHash ?? string.Empty,
            HashMode = string.IsNullOrWhiteSpace(revisionHash) ? string.Empty : HashMode,
            SnapshotPath = registration.LastSnapshotPath ?? string.Empty,
            SnapshotAtUtc = registration.LastSnapshotAtUtc ?? string.Empty,
            RecordedBy = recordedBy ?? string.Empty
        };
    }

    private static void ApplyManifestToState(FamilyBrowserStandardRevisionState state, FamilyBrowserStandardRevisionManifest manifest)
    {
        if (state == null || manifest == null)
        {
            return;
        }
        state.BaselineAtUtc = manifest.BaselineAtUtc ?? string.Empty;
        state.RecordedLastWriteUtc = manifest.SourceFileLastWriteUtc ?? string.Empty;
        state.RecordedLength = manifest.SourceFileLength;
        state.RecordedRevisionHash = manifest.RevisionHash ?? string.Empty;
        state.HashMode = manifest.HashMode ?? string.Empty;
        state.SnapshotPath = manifest.SnapshotPath ?? string.Empty;
        state.SnapshotAtUtc = manifest.SnapshotAtUtc ?? string.Empty;
    }

    private static FamilyBrowserStandardRevisionManifest LoadManifest(string workspaceRoot, string sourceId)
    {
        string path = GetManifestPath(workspaceRoot, sourceId);
        if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
        {
            return null;
        }
        lock (SyncRoot)
        {
            using (FileStream stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete))
            {
                DataContractJsonSerializer serializer = new DataContractJsonSerializer(typeof(FamilyBrowserStandardRevisionManifest));
                FamilyBrowserStandardRevisionManifest manifest = serializer.ReadObject(stream) as FamilyBrowserStandardRevisionManifest;
                return manifest != null && manifest.SchemaVersion == SchemaVersion ? manifest : null;
            }
        }
    }

    private static void SaveManifest(string workspaceRoot, FamilyBrowserStandardRevisionManifest manifest)
    {
        if (manifest == null || string.IsNullOrWhiteSpace(manifest.SourceId) || !FamilyBrowserStandardPolicyStore.IsManagedDataRootAvailable(workspaceRoot))
        {
            return;
        }
        string path = GetManifestPath(workspaceRoot, manifest.SourceId);
        lock (SyncRoot)
        {
            Directory.CreateDirectory(Path.GetDirectoryName(path));
            string temporary = FamilyBrowserAtomicFileService.CreateSiblingTemporaryPath(path);
            try
            {
                using (FileStream stream = new FileStream(temporary, FileMode.CreateNew, FileAccess.Write, FileShare.None))
                {
                    DataContractJsonSerializer serializer = new DataContractJsonSerializer(typeof(FamilyBrowserStandardRevisionManifest));
                    serializer.WriteObject(stream, manifest);
                    stream.Flush(true);
                }
                FamilyBrowserAtomicFileService.Promote(temporary, path);
            }
            finally
            {
                if (File.Exists(temporary))
                {
                    try { File.Delete(temporary); } catch { }
                }
            }
        }
    }

    private static string ResolveSourcePath(StandardLibraryRegistrationRecord registration)
    {
        if (registration == null)
        {
            return string.Empty;
        }
        return FamilyBrowserPathIdentityService.NormalizePath(!string.IsNullOrWhiteSpace(registration.ResolvedPath) ? registration.ResolvedPath : registration.Locator);
    }

    private static bool TryParseUtc(string value, out DateTime result)
    {
        return DateTime.TryParse(value ?? string.Empty, CultureInfo.InvariantCulture, DateTimeStyles.AssumeUniversal | DateTimeStyles.AdjustToUniversal, out result);
    }

    private static bool PathsEqual(string left, string right)
    {
        return string.Equals(FamilyBrowserPathIdentityService.NormalizePath(left), FamilyBrowserPathIdentityService.NormalizePath(right), StringComparison.OrdinalIgnoreCase);
    }

    private static void CopyRange(FileStream input, Stream output, long offset, long count)
    {
        input.Position = Math.Min(Math.Max(0L, offset), input.Length);
        byte[] buffer = new byte[64 * 1024];
        long remaining = Math.Min(count, input.Length - input.Position);
        while (remaining > 0L)
        {
            int read = input.Read(buffer, 0, (int)Math.Min(buffer.Length, remaining));
            if (read <= 0)
            {
                break;
            }
            output.Write(buffer, 0, read);
            remaining -= read;
        }
    }

    private static string SafeFileName(string value)
    {
        char[] invalid = Path.GetInvalidFileNameChars();
        return new string((value ?? string.Empty).Select(delegate(char ch) { return invalid.Contains(ch) ? '_' : ch; }).ToArray());
    }
}
