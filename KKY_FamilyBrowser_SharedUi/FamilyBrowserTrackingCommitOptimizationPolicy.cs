public enum FamilyBrowserPostCommitBaselineMode
{
    Incremental,
    FullCapture
}

public static class FamilyBrowserTrackingCommitOptimizationPolicy
{
    public static FamilyBrowserPostCommitBaselineMode ResolveBaselineMode(
        bool incrementalRefreshSucceeded,
        bool externalRebaseFailed,
        int eventReadFailureCount,
        int commitBoundaryReadFailureCount)
    {
        return incrementalRefreshSucceeded &&
            !externalRebaseFailed &&
            eventReadFailureCount == 0 &&
            commitBoundaryReadFailureCount == 0
            ? FamilyBrowserPostCommitBaselineMode.Incremental
            : FamilyBrowserPostCommitBaselineMode.FullCapture;
    }

    public static bool ShouldObserveProjectCatalog(
        bool decisionAvailable,
        bool catalogRelevantChange,
        bool evidenceGap)
    {
        return !decisionAvailable || catalogRelevantChange || evidenceGap;
    }
}
