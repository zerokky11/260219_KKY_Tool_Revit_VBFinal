using System;
using System.Collections.Generic;
using System.Linq;

public sealed class FamilyBrowserElementActivityMatchInput
{
    public List<string> ElementIds { get; set; }
    public List<string> TransactionNames { get; set; }

    public FamilyBrowserElementActivityMatchInput()
    {
        ElementIds = new List<string>();
        TransactionNames = new List<string>();
    }
}

public sealed class FamilyBrowserElementActivityMatchResult
{
    public List<int> CandidateIndexes { get; set; }
    public bool Exact { get; set; }

    public FamilyBrowserElementActivityMatchResult()
    {
        CandidateIndexes = new List<int>();
    }
}

public static class FamilyBrowserElementActivityMatcher
{
    public static FamilyBrowserElementActivityMatchResult Match(
        IList<FamilyBrowserElementActivityMatchInput> candidates,
        FamilyBrowserElementActivityMatchInput observed)
    {
        FamilyBrowserElementActivityMatchResult result = new FamilyBrowserElementActivityMatchResult();
        if (candidates == null || candidates.Count == 0 || observed == null)
        {
            return result;
        }

        HashSet<string> observedIds = ToSet(observed.ElementIds, StringComparer.Ordinal);
        HashSet<string> observedNames = ToSet(observed.TransactionNames, StringComparer.OrdinalIgnoreCase);
        if (observedIds.Count == 0)
        {
            return result;
        }

        HashSet<string> accumulatedIds = new HashSet<string>(StringComparer.Ordinal);
        HashSet<string> accumulatedNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        List<int> suffixIndexes = new List<int>();
        List<int> ambiguousSuffixIndexes = new List<int>();
        bool equivalentShorterSuffixExists = false;
        for (int i = candidates.Count - 1; i >= 0; i--)
        {
            FamilyBrowserElementActivityMatchInput candidate = candidates[i] ?? new FamilyBrowserElementActivityMatchInput();
            HashSet<string> candidateIds = ToSet(candidate.ElementIds, StringComparer.Ordinal);
            if (candidateIds.Any(delegate(string id) { return !observedIds.Contains(id); }))
            {
                break;
            }
            suffixIndexes.Add(i);
            accumulatedIds.UnionWith(candidateIds);
            accumulatedNames.UnionWith(ToSet(candidate.TransactionNames, StringComparer.OrdinalIgnoreCase));
            if (!accumulatedIds.SetEquals(observedIds) || !NamesMatch(accumulatedNames, observedNames))
            {
                continue;
            }

            if (ambiguousSuffixIndexes.Count == 0)
            {
                ambiguousSuffixIndexes = new List<int>(suffixIndexes);
            }
            else
            {
                equivalentShorterSuffixExists = true;
            }

            if (CanExtendWithoutNewObservedEvidence(candidates, i - 1, observedIds, observedNames))
            {
                equivalentShorterSuffixExists = true;
                continue;
            }

            result.CandidateIndexes = new List<int>(suffixIndexes);
            result.Exact = !equivalentShorterSuffixExists;
            return result;
        }
        if (ambiguousSuffixIndexes.Count > 0)
        {
            result.CandidateIndexes = ambiguousSuffixIndexes;
            return result;
        }

        for (int i = candidates.Count - 1; i >= 0; i--)
        {
            FamilyBrowserElementActivityMatchInput candidate = candidates[i] ?? new FamilyBrowserElementActivityMatchInput();
            HashSet<string> candidateIds = ToSet(candidate.ElementIds, StringComparer.Ordinal);
            if (candidateIds.SetEquals(observedIds))
            {
                result.CandidateIndexes.Add(i);
                return result;
            }
        }
        for (int i = candidates.Count - 1; i >= 0; i--)
        {
            FamilyBrowserElementActivityMatchInput candidate = candidates[i] ?? new FamilyBrowserElementActivityMatchInput();
            if (ToSet(candidate.ElementIds, StringComparer.Ordinal).Overlaps(observedIds))
            {
                result.CandidateIndexes.Add(i);
                return result;
            }
        }
        for (int i = candidates.Count - 1; i >= 0; i--)
        {
            FamilyBrowserElementActivityMatchInput candidate = candidates[i] ?? new FamilyBrowserElementActivityMatchInput();
            if (ToSet(candidate.TransactionNames, StringComparer.OrdinalIgnoreCase).Overlaps(observedNames))
            {
                result.CandidateIndexes.Add(i);
                return result;
            }
        }
        return result;
    }

    private static bool NamesMatch(ISet<string> candidateNames, ISet<string> observedNames)
    {
        return observedNames == null || observedNames.Count == 0 || candidateNames.SetEquals(observedNames);
    }

    private static bool CanExtendWithoutNewObservedEvidence(
        IList<FamilyBrowserElementActivityMatchInput> candidates,
        int candidateIndex,
        ISet<string> observedIds,
        ISet<string> observedNames)
    {
        if (candidates == null || candidateIndex < 0 || candidateIndex >= candidates.Count)
        {
            return false;
        }
        FamilyBrowserElementActivityMatchInput candidate = candidates[candidateIndex] ?? new FamilyBrowserElementActivityMatchInput();
        HashSet<string> candidateIds = ToSet(candidate.ElementIds, StringComparer.Ordinal);
        if (candidateIds.Count == 0 || candidateIds.Any(delegate(string id) { return !observedIds.Contains(id); }))
        {
            return false;
        }
        if (observedNames == null || observedNames.Count == 0)
        {
            return true;
        }
        HashSet<string> candidateNames = ToSet(candidate.TransactionNames, StringComparer.OrdinalIgnoreCase);
        return candidateNames.Count == 0 || candidateNames.All(delegate(string name) { return observedNames.Contains(name); });
    }

    private static HashSet<string> ToSet(IEnumerable<string> values, IEqualityComparer<string> comparer)
    {
        return new HashSet<string>((values ?? Enumerable.Empty<string>()).Where(delegate(string value) { return !string.IsNullOrWhiteSpace(value); }), comparer);
    }
}
