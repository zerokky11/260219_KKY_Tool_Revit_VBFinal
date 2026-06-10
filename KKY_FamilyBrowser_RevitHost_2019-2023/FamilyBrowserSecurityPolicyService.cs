using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using Microsoft.VisualBasic.CompilerServices;

public sealed class FamilyBrowserSecurityPolicyService
{
	private sealed class FamilyBrowserFileGuardPermissionDecision
	{
		public bool HasDecision { get; set; }

		public bool Allowed { get; set; }

		public FamilyBrowserFileGuardPermissionDecision()
		{
			HasDecision = false;
			Allowed = true;
		}
	}

	[CompilerGenerated]
	internal sealed class _Closure_0024__17_002D0
	{
		public List<string> _0024VB_0024Local_contextPaths;

		public HashSet<string> _0024VB_0024Local_contextNames;

		public Func<string, bool> _0024I1;

		public Func<string, bool> _0024I3;

		public _Closure_0024__17_002D0(_Closure_0024__17_002D0 arg0)
		{
			if (arg0 != null)
			{
				_0024VB_0024Local_contextPaths = arg0._0024VB_0024Local_contextPaths;
				_0024VB_0024Local_contextNames = arg0._0024VB_0024Local_contextNames;
			}
		}

		[SpecialName]
		internal bool _Lambda_0024__1(string targetPath)
		{
			_Closure_0024__17_002D1 arg = default(_Closure_0024__17_002D1);
			_Closure_0024__17_002D1 CS_0024_003C_003E8__locals2 = new _Closure_0024__17_002D1(arg);
			CS_0024_003C_003E8__locals2._0024VB_0024Local_targetPath = targetPath;
			return _0024VB_0024Local_contextPaths.Any([SpecialName] (string contextPath) => SamePath(contextPath, CS_0024_003C_003E8__locals2._0024VB_0024Local_targetPath));
		}

		[SpecialName]
		internal bool _Lambda_0024__3(string targetName)
		{
			return _0024VB_0024Local_contextNames.Contains(NormalizeDetachedFileBase(targetName));
		}
	}

	[CompilerGenerated]
	internal sealed class _Closure_0024__17_002D1
	{
		public string _0024VB_0024Local_targetPath;

		public _Closure_0024__17_002D1(_Closure_0024__17_002D1 arg0)
		{
			if (arg0 != null)
			{
				_0024VB_0024Local_targetPath = arg0._0024VB_0024Local_targetPath;
			}
		}

		[SpecialName]
		internal bool _Lambda_0024__2(string contextPath)
		{
			return SamePath(contextPath, _0024VB_0024Local_targetPath);
		}
	}

	private static string _currentUserIdentityOverride = string.Empty;

	private FamilyBrowserSecurityPolicyService()
	{
	}

	public static void SetCurrentUserIdentityOverride(string userIdentity)
	{
		_currentUserIdentityOverride = (userIdentity ?? string.Empty).Trim();
	}

	public static string ResolveCurrentUserIdentity()
	{
		if (!string.IsNullOrWhiteSpace(_currentUserIdentityOverride))
		{
			return _currentUserIdentityOverride;
		}
		string userName = (Environment.UserName ?? string.Empty).Trim();
		string domainName = (Environment.UserDomainName ?? string.Empty).Trim();
		if (string.IsNullOrWhiteSpace(domainName))
		{
			return userName;
		}
		return domainName + "\\" + userName;
	}

	public static string ResolveRole(FamilyBrowserStandardPolicy policy, string currentUser)
	{
		return ResolveRole(policy, currentUser, null);
	}

	public static string ResolveRole(FamilyBrowserStandardPolicy policy, string currentUser, FamilyBrowserProjectPolicyContext context)
	{
		return ResolveRoleFromSecurity(ResolveEffectiveSecurity(policy, context), currentUser);
	}

	public static string ResolveRole(FamilyBrowserStandardPolicy policy, string currentUser, FamilyBrowserProjectPolicyContext context, FamilyBrowserPermissionExcelDiagnostic excelDiagnostic)
	{
		return ResolveRoleFromSecurity(ResolveEffectiveSecurity(policy, context), currentUser);
	}

	private static string ResolveRoleFromSecurity(FamilyBrowserSecurityPolicy security, string currentUser)
	{
		if (security == null)
		{
			security = FamilyBrowserSecurityPolicy.CreateDefault();
		}
		if (security == null || !HasConfiguredAdmin(security))
		{
			return "Admin";
		}
		if (MatchesAnyUser(security.ReadOnlyUsers, currentUser))
		{
			return "ReadOnly";
		}
		if (HasAdminProfileKeywords(security))
		{
			if (MatchesAnyAdminProfileKeyword(security.AdminProfileKeywords, currentUser))
			{
				return "Admin";
			}
		}
		else if (MatchesAnyAdminUser(security.AdminUsers, currentUser))
		{
			return "Admin";
		}
		if (security.AllowUnlistedUsersAsModelers)
		{
			return "Modeler";
		}
		return "ReadOnly";
	}

	public static bool Can(FamilyBrowserStandardPolicy policy, string currentUser, string permission)
	{
		return Can(policy, currentUser, permission, null);
	}

	public static bool Can(FamilyBrowserStandardPolicy policy, string currentUser, string permission, FamilyBrowserProjectPolicyContext context)
	{
		FamilyBrowserSecurityPolicy security = ResolveEffectiveSecurity(policy, context);
		string role = ResolveRole(policy, currentUser, context);
		if (string.IsNullOrWhiteSpace(role))
		{
			role = ResolveRoleFromSecurity(security, currentUser);
		}
		if (string.Equals(role, "Admin", StringComparison.OrdinalIgnoreCase))
		{
			return true;
		}
		FamilyBrowserFileGuardPermissionDecision fileGuardDecision = ResolveFileGuardPermission(policy, permission, context);
		if (fileGuardDecision != null && fileGuardDecision.HasDecision)
		{
			return fileGuardDecision.Allowed;
		}
		FamilyBrowserPermissionExcelDecision excelDecision = FamilyBrowserPermissionExcelPolicyService.ResolvePermission(policy, currentUser, permission, context);
		if (excelDecision != null && excelDecision.HasDecision)
		{
			return excelDecision.Allowed;
		}
		if (FamilyBrowserPermissionExcelPolicyService.IsNativeGuardPermission(permission))
		{
			return true;
		}
		if (string.Equals(role, "ReadOnly", StringComparison.OrdinalIgnoreCase))
		{
			switch (permission)
			{
			case "ManagePolicy":
			case "RegisterStandard":
			case "LoadFamilies":
			case "ApplySystemTypes":
			case "StampTracking":
			case "CreateRequest":
			case "SubmitRequest":
			case "ApproveRequest":
			case "EditFamilies":
			case "AddDeleteTypes":
				return false;
			}
		}
		switch (permission)
		{
		case "ManagePolicy":
		case "RegisterStandard":
			return false;
		case "ApproveRequest":
			return false;
		case "LoadFamilies":
			return security.AllowModelersToLoadFamilies;
		case "ApplySystemTypes":
			return security.AllowModelersToApplySystemTypes;
		case "StampTracking":
			return !string.Equals(role, "ReadOnly", StringComparison.OrdinalIgnoreCase);
		case "EditFamilies":
		case "AddDeleteTypes":
			return string.Equals(role, "Admin", StringComparison.OrdinalIgnoreCase);
		case "CreateRequest":
		case "SubmitRequest":
			return string.Equals(role, "Approver", StringComparison.OrdinalIgnoreCase) || (string.Equals(role, "Modeler", StringComparison.OrdinalIgnoreCase) && security.AllowModelersToSubmitRequests);
		default:
			return !string.Equals(role, "ReadOnly", StringComparison.OrdinalIgnoreCase);
		}
	}

	public static bool CanNativeGuard(FamilyBrowserStandardPolicy policy, string currentUser, string permission, FamilyBrowserProjectPolicyContext context, bool adminModeEnabled)
	{
		FamilyBrowserSecurityPolicy security = ResolveEffectiveSecurity(policy, context);
		string role = ResolveRole(policy, currentUser, context);
		if (string.IsNullOrWhiteSpace(role))
		{
			role = ResolveRoleFromSecurity(security, currentUser);
		}
		if (string.Equals(role, "Admin", StringComparison.OrdinalIgnoreCase) && adminModeEnabled)
		{
			return true;
		}
		string effectiveRole = role;
		if (string.Equals(effectiveRole, "Admin", StringComparison.OrdinalIgnoreCase))
		{
			effectiveRole = "Modeler";
		}
		FamilyBrowserFileGuardPermissionDecision fileGuardDecision = ResolveFileGuardPermission(policy, permission, context);
		if (fileGuardDecision != null && fileGuardDecision.HasDecision)
		{
			return fileGuardDecision.Allowed;
		}
		FamilyBrowserPermissionExcelDecision excelDecision = FamilyBrowserPermissionExcelPolicyService.ResolvePermission(policy, currentUser, permission, context);
		if (excelDecision != null && excelDecision.HasDecision)
		{
			return excelDecision.Allowed;
		}
		if (FamilyBrowserPermissionExcelPolicyService.IsNativeGuardPermission(permission))
		{
			return true;
		}
		if (string.Equals(effectiveRole, "ReadOnly", StringComparison.OrdinalIgnoreCase))
		{
			switch (permission)
			{
			case "ManagePolicy":
			case "RegisterStandard":
			case "LoadFamilies":
			case "ApplySystemTypes":
			case "StampTracking":
			case "CreateRequest":
			case "SubmitRequest":
			case "ApproveRequest":
			case "EditFamilies":
			case "AddDeleteTypes":
				return false;
			}
		}
		switch (permission)
		{
		case "ManagePolicy":
		case "RegisterStandard":
			return false;
		case "ApproveRequest":
			return false;
		case "LoadFamilies":
			return security.AllowModelersToLoadFamilies;
		case "ApplySystemTypes":
			return security.AllowModelersToApplySystemTypes;
		case "StampTracking":
			return !string.Equals(effectiveRole, "ReadOnly", StringComparison.OrdinalIgnoreCase);
		case "EditFamilies":
		case "AddDeleteTypes":
			return false;
		case "CreateRequest":
		case "SubmitRequest":
			return string.Equals(effectiveRole, "Approver", StringComparison.OrdinalIgnoreCase) || (string.Equals(effectiveRole, "Modeler", StringComparison.OrdinalIgnoreCase) && security.AllowModelersToSubmitRequests);
		default:
			return !string.Equals(effectiveRole, "ReadOnly", StringComparison.OrdinalIgnoreCase);
		}
	}

	public static bool Can(FamilyBrowserStandardPolicy policy, string currentUser, string permission, FamilyBrowserProjectPolicyContext context, FamilyBrowserPermissionExcelDiagnostic excelDiagnostic)
	{
		FamilyBrowserSecurityPolicy security = ResolveEffectiveSecurity(policy, context);
		string role = ResolveRoleFromSecurity(security, currentUser);
		if (string.IsNullOrWhiteSpace(role))
		{
			role = ResolveRoleFromSecurity(security, currentUser);
		}
		if (string.Equals(role, "Admin", StringComparison.OrdinalIgnoreCase))
		{
			return true;
		}
		FamilyBrowserFileGuardPermissionDecision fileGuardDecision = ResolveFileGuardPermission(policy, permission, context);
		if (fileGuardDecision != null && fileGuardDecision.HasDecision)
		{
			return fileGuardDecision.Allowed;
		}
		FamilyBrowserPermissionExcelDecision excelDecision = FamilyBrowserPermissionExcelPolicyService.ResolvePermissionFromDiagnostic(excelDiagnostic, permission);
		if (excelDecision != null && excelDecision.HasDecision)
		{
			return excelDecision.Allowed;
		}
		if (FamilyBrowserPermissionExcelPolicyService.IsNativeGuardPermission(permission))
		{
			return true;
		}
		if (string.Equals(role, "ReadOnly", StringComparison.OrdinalIgnoreCase))
		{
			switch (permission)
			{
			case "ManagePolicy":
			case "RegisterStandard":
			case "LoadFamilies":
			case "ApplySystemTypes":
			case "StampTracking":
			case "CreateRequest":
			case "SubmitRequest":
			case "ApproveRequest":
			case "EditFamilies":
			case "AddDeleteTypes":
				return false;
			}
		}
		switch (permission)
		{
		case "ManagePolicy":
		case "RegisterStandard":
			return false;
		case "ApproveRequest":
			return false;
		case "LoadFamilies":
			return security.AllowModelersToLoadFamilies;
		case "ApplySystemTypes":
			return security.AllowModelersToApplySystemTypes;
		case "StampTracking":
			return !string.Equals(role, "ReadOnly", StringComparison.OrdinalIgnoreCase);
		case "EditFamilies":
		case "AddDeleteTypes":
			return string.Equals(role, "Admin", StringComparison.OrdinalIgnoreCase);
		case "CreateRequest":
		case "SubmitRequest":
			return string.Equals(role, "Approver", StringComparison.OrdinalIgnoreCase) || (string.Equals(role, "Modeler", StringComparison.OrdinalIgnoreCase) && security.AllowModelersToSubmitRequests);
		default:
			return !string.Equals(role, "ReadOnly", StringComparison.OrdinalIgnoreCase);
		}
	}

	private static string ResolveRoleFromDiagnostic(FamilyBrowserPermissionExcelDiagnostic diagnostic)
	{
		if (diagnostic == null || !diagnostic.Matched)
		{
			return string.Empty;
		}
		string left = FamilyBrowserPolicyKey.Normalize(diagnostic.MatchedRole);
		if (Operators.CompareString(left, FamilyBrowserPolicyKey.Normalize("Admin"), TextCompare: false) == 0 || Operators.CompareString(left, "admin", TextCompare: false) == 0 || Operators.CompareString(left, "administrator", TextCompare: false) == 0)
		{
			return "Admin";
		}
		if (Operators.CompareString(left, FamilyBrowserPolicyKey.Normalize("Approver"), TextCompare: false) == 0 || Operators.CompareString(left, "approver", TextCompare: false) == 0 || Operators.CompareString(left, "requestapprover", TextCompare: false) == 0 || Operators.CompareString(left, "requestapprover", TextCompare: false) == 0)
		{
			return "Approver";
		}
		if (Operators.CompareString(left, FamilyBrowserPolicyKey.Normalize("ReadOnly"), TextCompare: false) == 0 || Operators.CompareString(left, "readonly", TextCompare: false) == 0 || Operators.CompareString(left, "viewer", TextCompare: false) == 0)
		{
			return "ReadOnly";
		}
		if (Operators.CompareString(left, FamilyBrowserPolicyKey.Normalize("Modeler"), TextCompare: false) == 0 || Operators.CompareString(left, "modeler", TextCompare: false) == 0)
		{
			return "Modeler";
		}
		return string.Empty;
	}

	public static FamilyBrowserSecurityPolicy ResolveEffectiveSecurity(FamilyBrowserStandardPolicy policy, FamilyBrowserProjectPolicyContext context)
	{
		FamilyBrowserSecurityPolicy security = CloneSecurity((policy == null || policy.Security == null) ? FamilyBrowserSecurityPolicy.CreateDefault() : policy.Security);
		FamilyBrowserProjectPolicyRule rule = GetMatchingProjectRule(policy, context);
		if (rule == null)
		{
			return security;
		}
		ApplyProjectPreset(security, rule.PermissionPreset);
		List<string> targetUsers = security.AdminUsers;
		ApplyUserListOverride(ref targetUsers, rule.CustomAdminUsers);
		security.AdminUsers = targetUsers;
		targetUsers = security.RequestApproverUsers;
		ApplyUserListOverride(ref targetUsers, rule.CustomRequestApproverUsers);
		security.RequestApproverUsers = targetUsers;
		targetUsers = security.ReadOnlyUsers;
		ApplyUserListOverride(ref targetUsers, rule.CustomReadOnlyUsers);
		security.ReadOnlyUsers = targetUsers;
		bool targetValue = security.AllowUnlistedUsersAsModelers;
		ApplyBooleanOverride(ref targetValue, rule.AllowUnlistedUsersAsModelers);
		security.AllowUnlistedUsersAsModelers = targetValue;
		targetValue = security.AllowModelersToLoadFamilies;
		ApplyBooleanOverride(ref targetValue, rule.AllowModelersToLoadFamilies);
		security.AllowModelersToLoadFamilies = targetValue;
		targetValue = security.AllowModelersToApplySystemTypes;
		ApplyBooleanOverride(ref targetValue, rule.AllowModelersToApplySystemTypes);
		security.AllowModelersToApplySystemTypes = targetValue;
		targetValue = security.AllowModelersToSubmitRequests;
		ApplyBooleanOverride(ref targetValue, rule.AllowModelersToSubmitRequests);
		security.AllowModelersToSubmitRequests = targetValue;
		return security;
	}

	public static FamilyBrowserProjectPolicyRule GetMatchingProjectRule(FamilyBrowserStandardPolicy policy, FamilyBrowserProjectPolicyContext context)
	{
		if (policy == null || policy.ProjectPolicyRules == null || context == null)
		{
			return null;
		}
		return policy.ProjectPolicyRules.Where([SpecialName] (FamilyBrowserProjectPolicyRule x) => x?.Enabled ?? false).FirstOrDefault([SpecialName] (FamilyBrowserProjectPolicyRule x) => MatchesProjectRule(x, context));
	}

	private static FamilyBrowserFileGuardPermissionDecision ResolveFileGuardPermission(FamilyBrowserStandardPolicy policy, string permission, FamilyBrowserProjectPolicyContext context)
	{
		FamilyBrowserFileGuardPermissionDecision decision = new FamilyBrowserFileGuardPermissionDecision();
		if (!FamilyBrowserPermissionExcelPolicyService.IsNativeGuardPermission(permission))
		{
			return decision;
		}
		FamilyBrowserFileGuardPolicy fileGuard = policy?.FileGuard;
		if (fileGuard == null || !fileGuard.Enabled)
		{
			return decision;
		}
		decision.HasDecision = true;
		decision.Allowed = true;
		FamilyBrowserFileGuardTarget target = FindMatchingFileGuardTarget(fileGuard, context);
		if (target == null)
		{
			return decision;
		}
		if (string.Equals(permission, "EditFamilies", StringComparison.OrdinalIgnoreCase))
		{
			decision.Allowed = !target.BlockFamilyLoadAndEdit;
		}
		else if (string.Equals(permission, "AddDeleteTypes", StringComparison.OrdinalIgnoreCase))
		{
			decision.Allowed = !target.BlockTypeChanges;
		}
		return decision;
	}

	private static FamilyBrowserFileGuardTarget FindMatchingFileGuardTarget(FamilyBrowserFileGuardPolicy fileGuard, FamilyBrowserProjectPolicyContext context)
	{
		_Closure_0024__17_002D0 arg = default(_Closure_0024__17_002D0);
		_Closure_0024__17_002D0 CS_0024_003C_003E8__locals6 = new _Closure_0024__17_002D0(arg);
		if (fileGuard == null || fileGuard.Targets == null || context == null)
		{
			return null;
		}
		List<FamilyBrowserFileGuardTarget> enabledTargets = fileGuard.Targets.Where([SpecialName] (FamilyBrowserFileGuardTarget x) => x?.Enabled ?? false).ToList();
		if (enabledTargets.Count == 0)
		{
			return null;
		}
		CS_0024_003C_003E8__locals6._0024VB_0024Local_contextPaths = BuildContextPathCandidates(context);
		foreach (FamilyBrowserFileGuardTarget target in enabledTargets)
		{
			if (BuildTargetPathCandidates(fileGuard, target).Any([SpecialName] (string targetPath) =>
			{
				_Closure_0024__17_002D1 arg2 = default(_Closure_0024__17_002D1);
				_Closure_0024__17_002D1 CS_0024_003C_003E8__locals7 = new _Closure_0024__17_002D1(arg2);
				CS_0024_003C_003E8__locals7._0024VB_0024Local_targetPath = targetPath;
				return CS_0024_003C_003E8__locals6._0024VB_0024Local_contextPaths.Any([SpecialName] (string contextPath) => SamePath(contextPath, CS_0024_003C_003E8__locals7._0024VB_0024Local_targetPath));
			}))
			{
				return target;
			}
		}
		CS_0024_003C_003E8__locals6._0024VB_0024Local_contextNames = BuildContextFileNameCandidates(context);
		foreach (FamilyBrowserFileGuardTarget target2 in enabledTargets)
		{
			if (BuildTargetFileNameCandidates(target2).Any([SpecialName] (string targetName) => CS_0024_003C_003E8__locals6._0024VB_0024Local_contextNames.Contains(NormalizeDetachedFileBase(targetName))))
			{
				return target2;
			}
		}
		return null;
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

	private static List<string> BuildTargetPathCandidates(FamilyBrowserFileGuardPolicy fileGuard, FamilyBrowserFileGuardTarget target)
	{
		List<string> result = new List<string>();
		if (target == null)
		{
			return result;
		}
		AddPathCandidate(result, target.CentralPath);
		if (fileGuard != null && !string.IsNullOrWhiteSpace(fileGuard.RootFolder) && !string.IsNullOrWhiteSpace(target.RelativePath))
		{
			try
			{
				AddPathCandidate(result, Path.Combine(fileGuard.RootFolder, target.RelativePath));
			}
			catch (Exception projectError)
			{
				ProjectData.SetProjectError(projectError);
				ProjectData.ClearProjectError();
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
		AddNameCandidate(result, target.FileName);
		AddNameCandidate(result, target.CentralPath);
		AddNameCandidate(result, target.RelativePath);
		return result;
	}

	private static void AddPathCandidate(List<string> values, string value)
	{
		if (values != null && !string.IsNullOrWhiteSpace(value))
		{
			values.Add(value.Trim());
		}
	}

	private static void AddNameCandidate(HashSet<string> values, string value)
	{
		if (values != null)
		{
			string normalized = NormalizeDetachedFileBase(value);
			if (!string.IsNullOrWhiteSpace(normalized))
			{
				values.Add(normalized);
			}
		}
	}

	private static void AddNameCandidate(List<string> values, string value)
	{
		if (values != null && !string.IsNullOrWhiteSpace(value))
		{
			values.Add(value);
		}
	}

	private static bool SamePath(string leftValue, string rightValue)
	{
		return string.Equals(NormalizePathForCompare(leftValue), NormalizePathForCompare(rightValue), StringComparison.OrdinalIgnoreCase);
	}

	private static string NormalizePathForCompare(string value)
	{
		string text = (value ?? string.Empty).Trim();
		if (text.Length == 0)
		{
			return string.Empty;
		}
		text = text.Replace('/', '\\').TrimEnd('\\');
		try
		{
			if (Path.IsPathRooted(text))
			{
				text = Path.GetFullPath(text);
			}
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ProjectData.ClearProjectError();
		}
		return text.TrimEnd('\\');
	}

	private static string NormalizeDetachedFileBase(string value)
	{
		string text = (value ?? string.Empty).Trim();
		if (text.Length == 0)
		{
			return string.Empty;
		}
		checked
		{
			try
			{
				text = Path.GetFileNameWithoutExtension(text);
			}
			catch (Exception projectError)
			{
				ProjectData.SetProjectError(projectError);
				if (text.EndsWith(".rvt", StringComparison.OrdinalIgnoreCase))
				{
					text = text.Substring(0, text.Length - 4);
				}
				ProjectData.ClearProjectError();
			}
			text = text.Trim();
			string[] suffixes = new string[12]
			{
				"_detached", "-detached", ".detached", " detached", "(detached)", " - detached", " _ detached", "_detached copy", "_detached_copy", "-detached copy",
				"-detached-copy", " detached copy"
			};
			bool changed = true;
			while (changed && text.Length > 0)
			{
				changed = false;
				string[] array = suffixes;
				foreach (string suffix in array)
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
	}

	public static List<string> ParseUserList(string rawUsers)
	{
		char[] separators = new char[5] { ',', ';', '\r', '\n', '\t' };
		return (from x in (rawUsers ?? string.Empty).Split(separators, StringSplitOptions.RemoveEmptyEntries)
			select x.Trim() into x
			where x.Length > 0
			select x).Distinct(StringComparer.OrdinalIgnoreCase).OrderBy([SpecialName] (string x) => x, StringComparer.OrdinalIgnoreCase).ToList();
	}

	public static string FormatUserList(IEnumerable<string> users)
	{
		if (users == null)
		{
			return string.Empty;
		}
		return string.Join("; ", users.Where([SpecialName] (string x) => !string.IsNullOrWhiteSpace(x)));
	}

	public static bool HasConfiguredAdmin(FamilyBrowserSecurityPolicy security)
	{
		if (security != null)
		{
			if (security.AdminUsers == null || !security.AdminUsers.Any([SpecialName] (string x) => !string.IsNullOrWhiteSpace(x)))
			{
				return HasAdminProfileKeywords(security);
			}
			return true;
		}
		return false;
	}

	private static bool HasAdminProfileKeywords(FamilyBrowserSecurityPolicy security)
	{
		if (security != null && security.AdminProfileKeywords != null)
		{
			return security.AdminProfileKeywords.Any([SpecialName] (string x) => !string.IsNullOrWhiteSpace(x));
		}
		return false;
	}

	private static FamilyBrowserSecurityPolicy CloneSecurity(FamilyBrowserSecurityPolicy security)
	{
		if (security == null)
		{
			return FamilyBrowserSecurityPolicy.CreateDefault();
		}
		return new FamilyBrowserSecurityPolicy
		{
			AdminUsers = ParseUserList(FormatUserList(security.AdminUsers)),
			AdminProfileKeywords = ParseUserList(FormatUserList(security.AdminProfileKeywords)),
			RequestApproverUsers = ParseUserList(FormatUserList(security.RequestApproverUsers)),
			ReadOnlyUsers = ParseUserList(FormatUserList(security.ReadOnlyUsers)),
			AllowUnlistedUsersAsModelers = security.AllowUnlistedUsersAsModelers,
			AllowModelersToLoadFamilies = security.AllowModelersToLoadFamilies,
			AllowModelersToApplySystemTypes = security.AllowModelersToApplySystemTypes,
			AllowModelersToSubmitRequests = security.AllowModelersToSubmitRequests,
			LastUpdatedUtc = security.LastUpdatedUtc,
			LastUpdatedBy = security.LastUpdatedBy
		};
	}

	private static void ApplyProjectPreset(FamilyBrowserSecurityPolicy security, string preset)
	{
		string left = FamilyBrowserPolicyKey.Normalize(preset);
		if (Operators.CompareString(left, FamilyBrowserPolicyKey.Normalize("StandardModeler"), TextCompare: false) == 0)
		{
			security.AllowUnlistedUsersAsModelers = true;
			security.AllowModelersToLoadFamilies = true;
			security.AllowModelersToApplySystemTypes = true;
			security.AllowModelersToSubmitRequests = true;
		}
		else if (Operators.CompareString(left, FamilyBrowserPolicyKey.Normalize("RequestOnly"), TextCompare: false) == 0)
		{
			security.AllowUnlistedUsersAsModelers = true;
			security.AllowModelersToLoadFamilies = false;
			security.AllowModelersToApplySystemTypes = false;
			security.AllowModelersToSubmitRequests = true;
		}
		else if (Operators.CompareString(left, FamilyBrowserPolicyKey.Normalize("ReadOnly"), TextCompare: false) == 0)
		{
			security.AllowUnlistedUsersAsModelers = false;
			security.AllowModelersToLoadFamilies = false;
			security.AllowModelersToApplySystemTypes = false;
			security.AllowModelersToSubmitRequests = false;
		}
	}

	private static void ApplyUserListOverride(ref List<string> targetUsers, List<string> overrideUsers)
	{
		if (overrideUsers != null && overrideUsers.Count != 0)
		{
			targetUsers = ParseUserList(FormatUserList(overrideUsers));
		}
	}

	private static void ApplyBooleanOverride(ref bool targetValue, string overrideValue)
	{
		string left = FamilyBrowserPolicyKey.Normalize(overrideValue);
		if (Operators.CompareString(left, FamilyBrowserPolicyKey.Normalize("Allow"), TextCompare: false) == 0)
		{
			targetValue = true;
		}
		else if (Operators.CompareString(left, FamilyBrowserPolicyKey.Normalize("Deny"), TextCompare: false) == 0)
		{
			targetValue = false;
		}
	}

	private static bool MatchesProjectRule(FamilyBrowserProjectPolicyRule rule, FamilyBrowserProjectPolicyContext context)
	{
		if (rule == null || context == null)
		{
			return false;
		}
		string matchValue = (rule.MatchValue ?? string.Empty).Trim();
		string left = FamilyBrowserPolicyKey.Normalize(rule.MatchMode);
		if (Operators.CompareString(left, FamilyBrowserPolicyKey.Normalize("Any"), TextCompare: false) == 0)
		{
			return true;
		}
		if (Operators.CompareString(left, FamilyBrowserPolicyKey.Normalize("ExactCentralPath"), TextCompare: false) == 0)
		{
			return SameText(context.CentralPath, matchValue);
		}
		if (Operators.CompareString(left, FamilyBrowserPolicyKey.Normalize("ExactModelPath"), TextCompare: false) == 0)
		{
			return SameText(context.ModelPath, matchValue);
		}
		if (Operators.CompareString(left, FamilyBrowserPolicyKey.Normalize("ModelPathContains"), TextCompare: false) == 0)
		{
			return ContainsText(context.ModelPath, matchValue);
		}
		if (Operators.CompareString(left, FamilyBrowserPolicyKey.Normalize("ProjectTitleContains"), TextCompare: false) == 0)
		{
			return ContainsText(context.ProjectTitle, matchValue);
		}
		return ContainsText(context.CentralPath, matchValue);
	}

	private static bool SameText(string leftValue, string rightValue)
	{
		return string.Equals((leftValue ?? string.Empty).Trim(), (rightValue ?? string.Empty).Trim(), StringComparison.OrdinalIgnoreCase);
	}

	private static bool ContainsText(string value, string token)
	{
		if (string.IsNullOrWhiteSpace(token))
		{
			return false;
		}
		return (value ?? string.Empty).IndexOf(token.Trim(), StringComparison.OrdinalIgnoreCase) >= 0;
	}

	private static bool MatchesAnyUser(IEnumerable<string> users, string currentUser)
	{
		if (users == null)
		{
			return false;
		}
		HashSet<string> hashSet = BuildCurrentUserCandidates(currentUser);
		return users.Any([SpecialName] (string x) => hashSet.Contains(NormalizeUser(x)));
	}

	private static bool MatchesAnyAdminUser(IEnumerable<string> users, string currentUser)
	{
		if (MatchesAnyUser(users, currentUser))
		{
			return true;
		}
		if (users == null)
		{
			return false;
		}
		HashSet<string> currentUserCandidates = BuildCurrentUserCandidates(currentUser);
		return users.Any([SpecialName] (string x) => AdminKeywordMatchesCurrentUser(x, currentUserCandidates));
	}

	private static bool MatchesAnyAdminProfileKeyword(IEnumerable<string> keywords, string currentUser)
	{
		if (keywords == null)
		{
			return false;
		}
		HashSet<string> hashSet = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
		AddCandidate(hashSet, currentUser);
		if (!string.IsNullOrWhiteSpace(currentUser) && currentUser.Contains("\\"))
		{
			AddCandidate(hashSet, currentUser.Substring(checked(currentUser.LastIndexOf('\\') + 1)));
		}
		return keywords.Any([SpecialName] (string x) => AdminKeywordMatchesCurrentUser(x, hashSet));
	}

	private static bool AdminKeywordMatchesCurrentUser(string rawKeyword, HashSet<string> currentUserCandidates)
	{
		if (currentUserCandidates == null || currentUserCandidates.Count == 0)
		{
			return false;
		}
		string text = NormalizeUser(rawKeyword);
		if (string.IsNullOrWhiteSpace(text) || text.Contains("\\"))
		{
			return false;
		}
		if (text.StartsWith("contains:", StringComparison.OrdinalIgnoreCase))
		{
			text = text.Substring("contains:".Length).Trim();
		}
		else if (text.StartsWith("keyword:", StringComparison.OrdinalIgnoreCase))
		{
			text = text.Substring("keyword:".Length).Trim();
		}
		text = text.Trim('*', ' ');
		if (text.Length < 2)
		{
			return false;
		}
		return currentUserCandidates.Any([SpecialName] (string x) => x.IndexOf(text, StringComparison.OrdinalIgnoreCase) >= 0);
	}

	private static HashSet<string> BuildCurrentUserCandidates(string currentUser)
	{
		HashSet<string> values = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
		AddCandidate(values, currentUser);
		AddCandidate(values, Environment.UserName);
		AddCandidate(values, Environment.UserDomainName + "\\" + Environment.UserName);
		if (!string.IsNullOrWhiteSpace(currentUser) && currentUser.Contains("\\"))
		{
			AddCandidate(values, currentUser.Substring(checked(currentUser.LastIndexOf('\\') + 1)));
		}
		return values;
	}

	private static void AddCandidate(HashSet<string> values, string value)
	{
		string normalized = NormalizeUser(value);
		if (!string.IsNullOrWhiteSpace(normalized))
		{
			values.Add(normalized);
		}
	}

	private static string NormalizeUser(string value)
	{
		if (value == null)
		{
			return string.Empty;
		}
		return value.Trim().Replace('/', '\\').ToLowerInvariant();
	}
}
