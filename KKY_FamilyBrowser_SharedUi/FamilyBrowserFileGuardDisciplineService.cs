using System;
using System.Collections.Generic;
using System.Linq;

public static class FamilyBrowserFileGuardDisciplineService
{
	public static string ResolveAssignedDiscipline(
		FamilyBrowserStandardPolicy policy,
		FamilyBrowserFileGuardTarget target,
		bool allowLegacyFallback = true)
	{
		FamilyBrowserStandardLibrarySlot slot = ResolveSlot(policy, target == null ? string.Empty : target.Discipline);
		if (slot != null)
		{
			return slot.Discipline ?? string.Empty;
		}
		if (!allowLegacyFallback || policy == null)
		{
			return string.Empty;
		}
		slot = FamilyBrowserStandardPolicyStore.GetEffectiveSlot(policy);
		return slot == null ? string.Empty : (slot.Discipline ?? string.Empty);
	}

	public static FamilyBrowserStandardLibrarySlot ResolveSlot(
		FamilyBrowserStandardPolicy policy,
		string disciplineOrSlot)
	{
		if (policy == null)
		{
			return null;
		}
		if (string.Equals(policy.Mode, "Integrated", StringComparison.OrdinalIgnoreCase))
		{
			return policy.IntegratedLibrary != null && policy.IntegratedLibrary.Enabled
				? policy.IntegratedLibrary
				: null;
		}
		string value = (disciplineOrSlot ?? string.Empty).Trim();
		if (value.Length == 0)
		{
			return null;
		}
		string resolved = FamilyBrowserStandardPolicyStore.ResolveDisciplineKey(policy, value);
		string valueKey = FamilyBrowserPolicyKey.Normalize(value);
		string resolvedKey = FamilyBrowserPolicyKey.Normalize(resolved);
		return FamilyBrowserStandardPolicyStore.GetDisciplineSlots(policy).FirstOrDefault(delegate(FamilyBrowserStandardLibrarySlot slot)
		{
			if (slot == null)
			{
				return false;
			}
			string disciplineKey = FamilyBrowserPolicyKey.Normalize(slot.Discipline);
			string displayKey = FamilyBrowserPolicyKey.Normalize(slot.DisplayName);
			string slotKey = FamilyBrowserPolicyKey.Normalize(slot.SlotKey);
			return string.Equals(disciplineKey, resolvedKey, StringComparison.OrdinalIgnoreCase) ||
				string.Equals(disciplineKey, valueKey, StringComparison.OrdinalIgnoreCase) ||
				string.Equals(displayKey, valueKey, StringComparison.OrdinalIgnoreCase) ||
				string.Equals(slotKey, valueKey, StringComparison.OrdinalIgnoreCase) ||
				string.Equals(slotKey, "discipline-" + valueKey, StringComparison.OrdinalIgnoreCase);
		});
	}

	public static List<FamilyBrowserStandardLibrarySlot> GetSelectableSlots(FamilyBrowserStandardPolicy policy)
	{
		if (policy == null)
		{
			return new List<FamilyBrowserStandardLibrarySlot>();
		}
		if (string.Equals(policy.Mode, "Integrated", StringComparison.OrdinalIgnoreCase))
		{
			return policy.IntegratedLibrary != null && policy.IntegratedLibrary.Enabled
				? new List<FamilyBrowserStandardLibrarySlot> { policy.IntegratedLibrary }
				: new List<FamilyBrowserStandardLibrarySlot>();
		}
		return FamilyBrowserStandardPolicyStore.GetDisciplineSlots(policy);
	}
}
