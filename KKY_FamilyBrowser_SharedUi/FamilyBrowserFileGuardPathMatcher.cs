using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;

public sealed class FamilyBrowserFileGuardMatchResult
{
    public FamilyBrowserFileGuardTarget Target { get; set; }
    public string MatchKind { get; set; }
    public bool Ambiguous { get; set; }
    public int CandidateCount { get; set; }
    public bool IdentityUncertain { get; set; }
    public string Reason { get; set; }

    public FamilyBrowserFileGuardMatchResult()
    {
        MatchKind = "None";
        Reason = string.Empty;
    }
}

public static class FamilyBrowserFileGuardPathMatcher
{
    private sealed class MappedDriveCacheEntry
    {
        public string RemoteRoot { get; set; }
        public DateTime CachedUtc { get; set; }
    }

    private const int ErrorMoreData = 234;
    private static readonly TimeSpan MappedDriveCacheLifetime = TimeSpan.FromSeconds(30.0);
    private static readonly object CacheSync = new object();
    private static readonly Dictionary<string, MappedDriveCacheEntry> MappedDriveRootCache =
        new Dictionary<string, MappedDriveCacheEntry>(StringComparer.OrdinalIgnoreCase);

    [DllImport("mpr.dll", CharSet = CharSet.Unicode)]
    private static extern int WNetGetConnection(string localName, StringBuilder remoteName, ref int length);

    public static FamilyBrowserFileGuardTarget FindMatchingTarget(
        FamilyBrowserFileGuardPolicy fileGuard,
        FamilyBrowserProjectPolicyContext context)
    {
        return Resolve(fileGuard, context).Target;
    }

    public static FamilyBrowserFileGuardMatchResult Resolve(
        FamilyBrowserFileGuardPolicy fileGuard,
        FamilyBrowserProjectPolicyContext context)
    {
        FamilyBrowserFileGuardMatchResult result = new FamilyBrowserFileGuardMatchResult();
        if (fileGuard == null || !fileGuard.Enabled || fileGuard.Targets == null || context == null)
        {
            return result;
        }

        List<FamilyBrowserFileGuardTarget> enabledTargets = fileGuard.Targets
            .Where(delegate(FamilyBrowserFileGuardTarget target)
            {
                return target != null && target.Enabled;
            })
            .ToList();
        if (enabledTargets.Count == 0)
        {
            return result;
        }

        List<string> contextPaths = BuildContextPathCandidates(context);
		bool hasCentralPath = HasUsablePathCandidate(context.CentralPath);
		bool identityUncertain = context.IsWorkshared && (!hasCentralPath || IsPathIdentityResolutionUncertain(context.CentralPath));
        List<FamilyBrowserFileGuardTarget> pathMatches = enabledTargets
            .Where(delegate(FamilyBrowserFileGuardTarget target)
            {
                return TargetMatchesAnyPath(fileGuard, target, contextPaths);
            })
            .ToList();
        if (pathMatches.Count == 1)
        {
            result.Target = pathMatches[0];
            result.MatchKind = "PathIdentity";
            result.CandidateCount = 1;
            return result;
        }
        if (pathMatches.Count > 1)
        {
            result.Target = BuildConservativeDuplicateTarget(pathMatches);
            result.MatchKind = "ConservativeDuplicatePathIdentity";
            result.Ambiguous = true;
            result.CandidateCount = pathMatches.Count;
            return result;
        }

        // A filename fallback is only a lineage hint for a workshared local or detached
        // document whose central path has not been resolved yet. Standalone same-name
        // files and contexts with a different resolved central path must not inherit a guard.
        if (!context.IsWorkshared || (HasUsablePathCandidate(context.CentralPath) && !identityUncertain))
        {
            return result;
        }

        HashSet<string> contextNames = BuildContextFileNameCandidates(context);
        List<FamilyBrowserFileGuardTarget> nameMatches = enabledTargets
            .Where(delegate(FamilyBrowserFileGuardTarget target)
            {
                return BuildTargetFileNameCandidates(target).Any(delegate(string targetName)
                {
                    string normalized = NormalizeDetachedFileBase(targetName);
                    return normalized.Length > 0 && contextNames.Contains(normalized);
                });
            })
            .ToList();
        if (nameMatches.Count == 1)
        {
            result.Target = nameMatches[0];
			result.MatchKind = identityUncertain ? "UniqueWorksharedNamePendingIdentity" : "UniqueWorksharedNameFallback";
            result.CandidateCount = 1;
			result.IdentityUncertain = identityUncertain;
			result.Reason = identityUncertain ? "The workshared path identity is temporarily unavailable; the unique registered RVT name is enforced conservatively." : string.Empty;
            return result;
        }
        if (nameMatches.Count > 1)
        {
			if (identityUncertain)
			{
				result.Target = BuildConservativeDuplicateTarget(nameMatches);
				result.MatchKind = "ConservativeAmbiguousWorksharedNamePendingIdentity";
				result.IdentityUncertain = true;
				result.Reason = "The workshared path identity is temporarily unavailable and multiple registered RVTs share this name; the strictest combined guard is enforced until identity recovers.";
			}
			else
			{
				result.MatchKind = "AmbiguousWorksharedName";
			}
            result.Ambiguous = true;
            result.CandidateCount = nameMatches.Count;
        }
        return result;
    }

    public static bool PathsReferToSameFile(string leftValue, string rightValue)
    {
        HashSet<string> leftKeys = BuildFastPathIdentityKeys(leftValue);
        if (leftKeys.Count == 0)
        {
            return false;
        }
        HashSet<string> rightKeys = BuildFastPathIdentityKeys(rightValue);
        if (rightKeys.Any(delegate(string key) { return leftKeys.Contains(key); }))
        {
            return true;
        }

        if (!SameFileName(leftValue, rightValue))
        {
            return false;
        }

        // File identity is the final fallback for DFS, hard-link, or other aliases that
        // cannot be proven from their visible path forms. It is intentionally attempted
        // only for equal RVT names after the inexpensive path checks fail.
        string leftIdentity = FamilyBrowserPathIdentityService.GetComparableIdentity(leftValue);
        string rightIdentity = FamilyBrowserPathIdentityService.GetComparableIdentity(rightValue);
        return leftIdentity.StartsWith("FILE:", StringComparison.OrdinalIgnoreCase) &&
            string.Equals(leftIdentity, rightIdentity, StringComparison.OrdinalIgnoreCase);
    }

    public static string BuildStablePolicyPathKey(string value)
    {
        string normalized = NormalizePathForCompare(value);
        if (normalized.Length == 0)
        {
            return string.Empty;
        }
        string fileIdentity = FamilyBrowserPathIdentityService.GetFileIdentity(normalized);
        if (!string.IsNullOrWhiteSpace(fileIdentity))
        {
            return "FILE|" + fileIdentity.ToUpperInvariant();
        }
        string expanded = TryExpandMappedDrivePath(normalized);
        return "PATH|" + NormalizePathForCompare(expanded).ToUpperInvariant();
    }

    public static string Describe(FamilyBrowserFileGuardMatchResult match)
    {
        if (match == null)
        {
            return "None";
        }
        string targetPath = string.Empty;
        if (match.Target != null)
        {
            targetPath = FirstNonEmpty(match.Target.CentralPath, match.Target.RelativePath, match.Target.FileName);
        }
        return string.Join(";", new string[]
        {
            "kind=" + (match.MatchKind ?? "None"),
            "ambiguous=" + match.Ambiguous.ToString(),
			"identityUncertain=" + match.IdentityUncertain.ToString(),
            "candidates=" + match.CandidateCount.ToString(CultureInfo.InvariantCulture),
			"target=" + targetPath,
			"reason=" + (match.Reason ?? string.Empty)
        });
    }

	public static FamilyBrowserFileGuardTarget MergeConservativeTargets(IEnumerable<FamilyBrowserFileGuardTarget> targets)
	{
		return BuildConservativeDuplicateTarget(targets);
	}

    private static bool TargetMatchesAnyPath(
        FamilyBrowserFileGuardPolicy fileGuard,
        FamilyBrowserFileGuardTarget target,
        List<string> contextPaths)
    {
        if (contextPaths == null || contextPaths.Count == 0)
        {
            return false;
        }
        foreach (string targetPath in BuildTargetPathCandidates(fileGuard, target))
        {
            foreach (string contextPath in contextPaths)
            {
                if (PathsReferToSameFile(contextPath, targetPath))
                {
                    return true;
                }
            }
        }
        return false;
    }

    private static List<string> BuildContextPathCandidates(FamilyBrowserProjectPolicyContext context)
    {
        List<string> result = new List<string>();
        if (context == null)
        {
            return result;
        }
        AddPathCandidate(result, context.CentralPath);
        AddPathCandidate(result, context.ModelPath);
        return result;
    }

    private static List<string> BuildTargetPathCandidates(
        FamilyBrowserFileGuardPolicy fileGuard,
        FamilyBrowserFileGuardTarget target)
    {
        List<string> result = new List<string>();
        if (target == null)
        {
            return result;
        }
        AddPathCandidate(result, target.CentralPath);
        if (fileGuard != null &&
            HasUsablePathCandidate(fileGuard.RootFolder) &&
            !string.IsNullOrWhiteSpace(target.RelativePath))
        {
            try
            {
                AddPathCandidate(result, Path.Combine(fileGuard.RootFolder, target.RelativePath));
            }
            catch
            {
            }
        }
        return result;
    }

    private static HashSet<string> BuildContextFileNameCandidates(FamilyBrowserProjectPolicyContext context)
    {
        HashSet<string> result = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        if (context == null)
        {
            return result;
        }
        AddNameCandidate(result, context.CentralPath);
        AddNameCandidate(result, context.ModelPath);
        AddNameCandidate(result, context.ProjectTitle);
        return result;
    }

    private static List<string> BuildTargetFileNameCandidates(FamilyBrowserFileGuardTarget target)
    {
        List<string> result = new List<string>();
        if (target == null)
        {
            return result;
        }
        result.Add(target.FileName ?? string.Empty);
        result.Add(target.CentralPath ?? string.Empty);
        result.Add(target.RelativePath ?? string.Empty);
        return result;
    }

    private static void AddPathCandidate(List<string> values, string value)
    {
        if (values == null || !HasUsablePathCandidate(value))
        {
            return;
        }
        string trimmed = value.Trim();
        if (!values.Contains(trimmed, StringComparer.OrdinalIgnoreCase))
        {
            values.Add(trimmed);
        }
    }

    private static bool HasUsablePathCandidate(string value)
    {
        string text = (value ?? string.Empty).Trim();
        if (text.Length == 0)
        {
            return false;
        }
        if (text.StartsWith("\\\\", StringComparison.Ordinal) || text.IndexOf("://", StringComparison.Ordinal) > 0)
        {
            return true;
        }
        try
        {
            return Path.IsPathRooted(text);
        }
        catch
        {
            return false;
        }
    }

    private static HashSet<string> BuildFastPathIdentityKeys(string value)
    {
        HashSet<string> result = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        string normalized = NormalizePathForCompare(value);
        AddPathIdentityKey(result, normalized);
        string expanded = TryExpandMappedDrivePath(normalized);
        if (!string.Equals(expanded, normalized, StringComparison.OrdinalIgnoreCase))
        {
            AddPathIdentityKey(result, expanded);
        }
        return result;
    }

    private static void AddPathIdentityKey(HashSet<string> result, string value)
    {
        if (result == null)
        {
            return;
        }
        string normalized = NormalizePathForCompare(value);
        if (normalized.Length == 0)
        {
            return;
        }
        result.Add("PATH|" + normalized);
    }

    private static string TryExpandMappedDrivePath(string value)
    {
        string normalized = NormalizePathForCompare(value);
        if (normalized.Length < 2 || normalized[1] != ':' || !char.IsLetter(normalized[0]))
        {
            return normalized;
        }
        string driveName = char.ToUpperInvariant(normalized[0]).ToString(CultureInfo.InvariantCulture) + ":";
        string remoteRoot = ResolveMappedDriveRoot(driveName);
        if (remoteRoot.Length == 0)
        {
            return normalized;
        }
        string remainder = normalized.Substring(2).TrimStart('\\');
        return NormalizePathForCompare(remainder.Length == 0
            ? remoteRoot
            : remoteRoot.TrimEnd('\\') + "\\" + remainder);
    }

	private static bool IsPathIdentityResolutionUncertain(string value)
	{
		string normalized = NormalizePathForCompare(value);
		if (normalized.Length == 0 || normalized.IndexOf("://", StringComparison.Ordinal) > 0)
		{
			return false;
		}
		try
		{
			if (File.Exists(normalized))
			{
				return false;
			}
		}
		catch
		{
		}
		if (normalized.StartsWith("\\\\", StringComparison.Ordinal))
		{
			return true;
		}
		if (normalized.Length >= 2 && normalized[1] == ':' && char.IsLetter(normalized[0]))
		{
			string expanded = TryExpandMappedDrivePath(normalized);
			if (!string.Equals(expanded, normalized, StringComparison.OrdinalIgnoreCase) && expanded.StartsWith("\\\\", StringComparison.Ordinal))
			{
				return true;
			}
			try
			{
				string root = Path.GetPathRoot(normalized);
				return !string.IsNullOrWhiteSpace(root) && new DriveInfo(root).DriveType == DriveType.Network;
			}
			catch
			{
				return false;
			}
		}
		return false;
	}

    private static string ResolveMappedDriveRoot(string driveName)
    {
        if (string.IsNullOrWhiteSpace(driveName))
        {
            return string.Empty;
        }
        DateTime now = DateTime.UtcNow;
        lock (CacheSync)
        {
            MappedDriveCacheEntry cached;
            if (MappedDriveRootCache.TryGetValue(driveName, out cached) &&
                cached != null &&
                now - cached.CachedUtc < MappedDriveCacheLifetime)
            {
                return cached.RemoteRoot ?? string.Empty;
            }
        }

        string resolved = string.Empty;
        try
        {
            int length = 512;
            StringBuilder remoteName = new StringBuilder(length);
            int status = WNetGetConnection(driveName, remoteName, ref length);
            if (status == ErrorMoreData && length > remoteName.Capacity)
            {
                remoteName = new StringBuilder(length);
                status = WNetGetConnection(driveName, remoteName, ref length);
            }
            if (status == 0)
            {
                resolved = NormalizePathForCompare(remoteName.ToString());
            }
        }
        catch
        {
            resolved = string.Empty;
        }

		if (!string.IsNullOrWhiteSpace(resolved))
        {
			lock (CacheSync)
            {
				MappedDriveRootCache[driveName] = new MappedDriveCacheEntry
				{
					RemoteRoot = resolved,
					CachedUtc = now
				};
			}
        }
        return resolved;
    }

    private static FamilyBrowserFileGuardTarget BuildConservativeDuplicateTarget(IEnumerable<FamilyBrowserFileGuardTarget> matches)
    {
        List<FamilyBrowserFileGuardTarget> targets = (matches ?? Enumerable.Empty<FamilyBrowserFileGuardTarget>())
            .Where(delegate(FamilyBrowserFileGuardTarget target) { return target != null; })
            .OrderByDescending(delegate(FamilyBrowserFileGuardTarget target) { return target.LastUpdatedUtc ?? string.Empty; }, StringComparer.Ordinal)
            .ToList();
        if (targets.Count == 0)
        {
            return null;
        }
        FamilyBrowserFileGuardTarget preferred = targets[0];
        List<string> disciplines = targets
            .Select(delegate(FamilyBrowserFileGuardTarget target) { return (target.Discipline ?? string.Empty).Trim(); })
            .Where(delegate(string discipline) { return discipline.Length > 0; })
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();
        return new FamilyBrowserFileGuardTarget
        {
            Enabled = targets.Any(delegate(FamilyBrowserFileGuardTarget target) { return target.Enabled; }),
            FileName = preferred.FileName ?? string.Empty,
            CentralPath = preferred.CentralPath ?? string.Empty,
            RelativePath = preferred.RelativePath ?? string.Empty,
            Discipline = disciplines.Count == 1 ? disciplines[0] : string.Empty,
            BlockFamilyLoadAndEdit = targets.Any(delegate(FamilyBrowserFileGuardTarget target) { return target.BlockFamilyLoadAndEdit; }),
            BlockTypeChanges = targets.Any(delegate(FamilyBrowserFileGuardTarget target) { return target.BlockTypeChanges; }),
            BlockNestedOnlyStandalonePlacement = targets.Any(delegate(FamilyBrowserFileGuardTarget target) { return target.BlockNestedOnlyStandalonePlacement; }),
            TrackElementChanges = targets.Any(delegate(FamilyBrowserFileGuardTarget target) { return target.TrackElementChanges; }),
            TrackElementChangesConfigured = targets.Any(delegate(FamilyBrowserFileGuardTarget target) { return target.TrackElementChangesConfigured; }),
            LastUpdatedUtc = preferred.LastUpdatedUtc ?? string.Empty,
            LastUpdatedBy = preferred.LastUpdatedBy ?? string.Empty
        };
    }

    private static string NormalizePathForCompare(string value)
    {
        string text = Environment.ExpandEnvironmentVariables((value ?? string.Empty).Trim());
        if (text.Length == 0)
        {
            return string.Empty;
        }
        text = text.Replace('/', '\\').TrimEnd('\\');
        if (text.StartsWith("\\\\?\\UNC\\", StringComparison.OrdinalIgnoreCase))
        {
            text = "\\\\" + text.Substring(8);
        }
        else if (text.StartsWith("\\\\?\\", StringComparison.OrdinalIgnoreCase))
        {
            text = text.Substring(4);
        }
        if (text.IndexOf("://", StringComparison.Ordinal) > 0)
        {
            return text.TrimEnd('\\');
        }
        try
        {
            if (Path.IsPathRooted(text))
            {
                text = Path.GetFullPath(text);
            }
        }
        catch
        {
        }
        return text.TrimEnd('\\');
    }

    private static bool SameFileName(string leftValue, string rightValue)
    {
        string leftName = NormalizeDetachedFileBase(leftValue);
        string rightName = NormalizeDetachedFileBase(rightValue);
        return leftName.Length > 0 && string.Equals(leftName, rightName, StringComparison.OrdinalIgnoreCase);
    }

    private static void AddNameCandidate(HashSet<string> values, string value)
    {
        if (values == null)
        {
            return;
        }
        string normalized = NormalizeDetachedFileBase(value);
        if (normalized.Length > 0)
        {
            values.Add(normalized);
        }
    }

    private static string NormalizeDetachedFileBase(string value)
    {
        string text = (value ?? string.Empty).Trim();
        if (text.Length == 0)
        {
            return string.Empty;
        }
        try
        {
            text = Path.GetFileNameWithoutExtension(text);
        }
        catch
        {
            if (text.EndsWith(".rvt", StringComparison.OrdinalIgnoreCase))
            {
                text = text.Substring(0, text.Length - 4);
            }
        }
        text = text.Trim();
        string[] suffixes = new string[]
        {
            "_detached", "-detached", ".detached", " detached", "(detached)",
            " - detached", " _ detached", "_detached copy", "_detached_copy",
            "-detached copy", "-detached-copy", " detached copy"
        };
        bool changed = true;
        while (changed && text.Length > 0)
        {
            changed = false;
            foreach (string suffix in suffixes)
            {
                if (text.EndsWith(suffix, StringComparison.OrdinalIgnoreCase))
                {
                    text = text.Substring(0, text.Length - suffix.Length).Trim();
                    changed = true;
                    break;
                }
            }
        }
        return text.ToLowerInvariant();
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
}
