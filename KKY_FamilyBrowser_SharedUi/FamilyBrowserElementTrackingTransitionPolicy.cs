using System;
using System.Collections.Generic;

public static class FamilyBrowserElementTrackingTransitionPolicy
{
    public static void RestoreVisibleElementId(
        string elementId,
        ISet<string> currentEventIgnoredElementIds,
        ISet<string> sessionIgnoredElementIds)
    {
        if (string.IsNullOrWhiteSpace(elementId))
        {
            return;
        }
        if (currentEventIgnoredElementIds != null)
        {
            currentEventIgnoredElementIds.Remove(elementId);
        }
        if (sessionIgnoredElementIds != null)
        {
            sessionIgnoredElementIds.Remove(elementId);
        }
    }

    public static bool ShouldIgnoreChangedElement(
        bool resolvedAsAuxiliary,
        bool liveElementUnavailable,
        bool ignoredInCurrentEvent,
        bool ignoredInSession)
    {
        return resolvedAsAuxiliary ||
            (liveElementUnavailable && (ignoredInCurrentEvent || ignoredInSession));
    }

    public static string ResolveChangeKind(
        bool hasBefore,
        bool hasAfter,
        bool elementSequenceAmbiguous,
        bool hasActiveActivity,
        bool stateSignatureChanged,
        bool wasAdded,
        bool wasDeleted,
        bool baselineCapturedLate)
    {
        if (!hasBefore && hasAfter)
        {
            return "Created";
        }
        if (hasBefore && !hasAfter)
        {
            return "Deleted";
        }
        if (hasBefore && hasAfter &&
            ((!elementSequenceAmbiguous && hasActiveActivity) || stateSignatureChanged))
        {
            return "Modified";
        }
        if (!hasBefore && !hasAfter && !elementSequenceAmbiguous && wasAdded && wasDeleted)
        {
            return "CreatedThenDeleted";
        }
        if (!hasBefore && !hasAfter && !elementSequenceAmbiguous && baselineCapturedLate && hasActiveActivity && wasDeleted)
        {
            return "Deleted";
        }
        return string.Empty;
    }

    public static bool IsUnresolvedTransient(string changeKind, bool hasBefore, bool hasAfter, bool hasLastKnown)
    {
        return string.Equals(changeKind, "CreatedThenDeleted", StringComparison.OrdinalIgnoreCase) &&
            !hasBefore && !hasAfter && !hasLastKnown;
    }
}
