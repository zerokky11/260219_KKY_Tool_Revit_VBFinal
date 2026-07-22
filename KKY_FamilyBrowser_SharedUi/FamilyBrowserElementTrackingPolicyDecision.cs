public enum FamilyBrowserElementTrackingSessionMode
{
    Ignore = 0,
    Live = 1,
    DeferredCommit = 2,
    RecoveryOnly = 3
}

public static class FamilyBrowserElementTrackingPolicyDecision
{
    public static FamilyBrowserElementTrackingSessionMode Resolve(
        bool policyEnabled,
        bool policyStateIsFallbackOrDeferred,
        bool sessionExists,
        bool hasUncommittedEvidence,
        bool hasProtectedRecoveryEvidence)
    {
        if (policyEnabled && !policyStateIsFallbackOrDeferred)
        {
            return FamilyBrowserElementTrackingSessionMode.Live;
        }
        if (sessionExists && hasUncommittedEvidence)
        {
            return FamilyBrowserElementTrackingSessionMode.DeferredCommit;
        }
        if (sessionExists && hasProtectedRecoveryEvidence)
        {
            return FamilyBrowserElementTrackingSessionMode.RecoveryOnly;
        }
        return FamilyBrowserElementTrackingSessionMode.Ignore;
    }
}
