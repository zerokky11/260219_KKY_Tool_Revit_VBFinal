using System;
using System.Collections.Generic;

public sealed class SystemTypeSupportPolicyService
{
	public const string AutoApply = "AutoApply";

	public const string AutoApplyWithDependencies = "AutoApplyWithDependencies";

	public const string PreflightThenConfirm = "PreflightThenConfirm";

	public const string ReviewOnly = "ReviewOnly";

	private static readonly Dictionary<string, string> Policies = new Dictionary<string, string>(StringComparer.Ordinal)
	{
		{ "walltype", "AutoApply" },
		{ "floortype", "AutoApply" },
		{ "rooftype", "AutoApply" },
		{ "ceilingtype", "AutoApply" },
		{ "ductinsulationtype", "AutoApply" },
		{ "pipeinsulationtype", "AutoApply" },
		{ "ductliningtype", "AutoApply" },
		{ "pipetype", "AutoApplyWithDependencies" },
		{ "ducttype", "AutoApplyWithDependencies" },
		{ "flexpipetype", "AutoApplyWithDependencies" },
		{ "flexducttype", "AutoApplyWithDependencies" },
		{ "cabletraytype", "AutoApplyWithDependencies" },
		{ "conduittype", "AutoApplyWithDependencies" },
		{ "wiretype", "AutoApply" },
		{ "pipingsystemtype", "PreflightThenConfirm" },
		{ "ductsystemtype", "PreflightThenConfirm" },
		{ "mechanicalsystemtype", "PreflightThenConfirm" },
		{ "electricalsystemtype", "PreflightThenConfirm" },
		{ "stairstype", "ReviewOnly" },
		{ "railingtype", "ReviewOnly" }
	};

	private SystemTypeSupportPolicyService()
	{
	}

	public static string ResolvePolicy(string systemFamilyKind)
	{
		string policy = null;
		if (Policies.TryGetValue(Normalize(systemFamilyKind), out policy))
		{
			return policy;
		}
		return "ReviewOnly";
	}

	public static bool CanApply(string systemFamilyKind)
	{
		switch (ResolvePolicy(systemFamilyKind))
		{
		case "AutoApply":
		case "AutoApplyWithDependencies":
		case "PreflightThenConfirm":
			return true;
		default:
			return false;
		}
	}

	public static bool RequiresDependencyRefresh(string systemFamilyKind)
	{
		return string.Equals(ResolvePolicy(systemFamilyKind), "AutoApplyWithDependencies", StringComparison.Ordinal);
	}

	public static string ResolveDisplayLabel(string systemFamilyKind)
	{
		return ResolvePolicy(systemFamilyKind) switch
		{
			"AutoApply" => "Auto apply", 
			"AutoApplyWithDependencies" => "Auto apply with dependencies", 
			"PreflightThenConfirm" => "Confirm then apply", 
			_ => "Review only", 
		};
	}

	public static string SupportedApplySummary()
	{
		return "Auto: Wall/Floor/Roof/Ceiling, insulation, lining and Wire. Dependencies: Pipe/Duct/Flex/CableTray/Conduit. Confirm: Piping/Duct/Mechanical/Electrical system types.";
	}

	private static string Normalize(string value)
	{
		if (value == null)
		{
			return string.Empty;
		}
		return value.Trim().ToLowerInvariant();
	}
}
