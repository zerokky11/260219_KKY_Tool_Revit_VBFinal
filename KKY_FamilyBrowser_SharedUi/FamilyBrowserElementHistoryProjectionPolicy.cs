using System;
using System.Collections.Generic;

public enum FamilyBrowserElementHistoryProjectionDisposition
{
    Visible,
    AuxiliarySupport,
    UnresolvedTransient
}

public sealed class FamilyBrowserElementHistoryProjectionCounts
{
    public int VisibleChangeCount { get; set; }
    public int CreatedCount { get; set; }
    public int ModifiedCount { get; set; }
    public int DeletedCount { get; set; }
    public int TransientCreatedDeletedCount { get; set; }
    public int HiddenAuxiliaryCount { get; set; }
    public int HiddenUnresolvedTransientCount { get; set; }
}

public static class FamilyBrowserElementHistoryProjectionPolicy
{
    public const string UnresolvedTransientElementClass = "UnresolvedTransient";
    public const string UnresolvedTransientTrackingKind = "UnresolvedTransient";

    public static FamilyBrowserElementHistoryProjectionDisposition Classify(FamilyBrowserElementChangeItem change)
    {
        if (change == null)
        {
            return FamilyBrowserElementHistoryProjectionDisposition.AuxiliarySupport;
        }

        FamilyBrowserTrackedElementState state = change.After ?? change.Before;
        if (IsUnresolvedTransient(change, state))
        {
            return FamilyBrowserElementHistoryProjectionDisposition.UnresolvedTransient;
        }

        bool isElementType = state != null
            ? state.IsElementType
            : (change.ElementClass ?? string.Empty).EndsWith("Type", StringComparison.OrdinalIgnoreCase);
        bool auxiliary = FamilyBrowserElementTrackingScopePolicy.IsAuxiliarySupportRecord(
            state == null ? change.ElementClass : state.ElementClass,
            state == null ? string.Empty : state.CategoryId,
            state == null ? change.CategoryName : state.CategoryName,
            isElementType);
        return auxiliary
            ? FamilyBrowserElementHistoryProjectionDisposition.AuxiliarySupport
            : FamilyBrowserElementHistoryProjectionDisposition.Visible;
    }

    public static bool IsUserFacingChange(FamilyBrowserElementChangeItem change)
    {
        return Classify(change) == FamilyBrowserElementHistoryProjectionDisposition.Visible;
    }

    public static FamilyBrowserElementHistoryProjectionCounts CountUserFacingChanges(IEnumerable<FamilyBrowserElementChangeItem> changes)
    {
        FamilyBrowserElementHistoryProjectionCounts counts = new FamilyBrowserElementHistoryProjectionCounts();
        foreach (FamilyBrowserElementChangeItem change in changes ?? new FamilyBrowserElementChangeItem[0])
        {
            FamilyBrowserElementHistoryProjectionDisposition disposition = Classify(change);
            if (disposition == FamilyBrowserElementHistoryProjectionDisposition.AuxiliarySupport)
            {
                counts.HiddenAuxiliaryCount++;
                continue;
            }
            if (disposition == FamilyBrowserElementHistoryProjectionDisposition.UnresolvedTransient)
            {
                counts.HiddenUnresolvedTransientCount++;
                continue;
            }

            counts.VisibleChangeCount++;
            string kind = change.ChangeKind ?? string.Empty;
            if (string.Equals(kind, "Created", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(kind, "CreatedThenDeleted", StringComparison.OrdinalIgnoreCase))
            {
                counts.CreatedCount++;
            }
            if (string.Equals(kind, "Modified", StringComparison.OrdinalIgnoreCase))
            {
                counts.ModifiedCount++;
            }
            if (string.Equals(kind, "Deleted", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(kind, "CreatedThenDeleted", StringComparison.OrdinalIgnoreCase))
            {
                counts.DeletedCount++;
            }
            if (string.Equals(kind, "CreatedThenDeleted", StringComparison.OrdinalIgnoreCase))
            {
                counts.TransientCreatedDeletedCount++;
            }
        }
        return counts;
    }

    private static bool IsUnresolvedTransient(FamilyBrowserElementChangeItem change, FamilyBrowserTrackedElementState state)
    {
        if (string.Equals(change.TrackingKind, UnresolvedTransientTrackingKind, StringComparison.OrdinalIgnoreCase) ||
            string.Equals(change.ElementClass, UnresolvedTransientElementClass, StringComparison.OrdinalIgnoreCase))
        {
            return true;
        }
        return state == null &&
            string.Equals(change.ChangeKind, "CreatedThenDeleted", StringComparison.OrdinalIgnoreCase) &&
            string.IsNullOrWhiteSpace(change.ElementClass) &&
            string.IsNullOrWhiteSpace(change.CategoryName) &&
            string.IsNullOrWhiteSpace(change.ElementName) &&
            string.IsNullOrWhiteSpace(change.FamilyName) &&
            string.IsNullOrWhiteSpace(change.TypeName);
    }
}
