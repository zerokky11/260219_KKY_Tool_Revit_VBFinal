using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using Autodesk.Revit.DB;

namespace KKY_Tool_Revit.Services
{
    public static class LegacyPurgeService
    {
        public static int Run(Document doc, int passCount, ElementId keepViewId, ElementId keepFilterId, Action<string> log)
        {
            if (doc == null) throw new ArgumentNullException(nameof(doc));

            int totalDeleted = 0;
            int totalPasses = Math.Max(1, passCount);
            for (int pass = 1; pass <= totalPasses; pass++)
            {
                PurgePassContext context = BuildPassContext(doc, keepViewId, keepFilterId);
                LogPurgeScope(doc, context, pass, log);

                int deletedThisPass = 0;
                deletedThisPass += TryNativePurge(doc, context, log);
                deletedThisPass += DeleteUnusedElementTypes(doc, context, log);
                deletedThisPass += DeleteUnusedTextNoteTypes(doc, context, log);
                deletedThisPass += DeleteUnusedFilters(doc, context, log);
                deletedThisPass += DeleteProbeSafeCandidates(doc, context, log);

                log?.Invoke($"Purge pass {pass}: 삭제 {deletedThisPass}개");
                totalDeleted += deletedThisPass;

                if (deletedThisPass == 0)
                {
                    log?.Invoke($"Purge pass {pass}: 더 이상 삭제 가능한 미사용 항목이 없어 종료합니다.");
                    break;
                }
            }

            log?.Invoke($"Purge 총 삭제: {totalDeleted}개");
            return totalDeleted;
        }

        private static PurgePassContext BuildPassContext(Document doc, ElementId keepViewId, ElementId keepFilterId)
        {
            var context = new PurgePassContext();
            context.KeepViewId = keepViewId;
            context.KeepFilterId = keepFilterId;
            context.UsedTypeIds = BuildUsedTypeIdSet(doc);
            context.UsedFilterIds = BuildUsedFilterIdSet(doc);
            context.ModelElementIdSet = BuildModelElementIdSet(doc);
            context.UnusedElementTypeCandidates = CollectUnusedElementTypeCandidates(doc, keepViewId, context.UsedTypeIds);
            context.ProbeCandidates = CollectGenericPurgeCandidates(doc, keepFilterId);
            context.SafeDeletedIdSet = BuildSafeDeletedIdSet(doc, keepFilterId);
            return context;
        }

        private static int TryNativePurge(Document doc, PurgePassContext context, Action<string> log)
        {
            MethodInfo method = typeof(Document).GetMethod("GetUnusedElements", new[] { typeof(ISet<ElementId>) });
            if (method == null)
            {
                log?.Invoke("Native purge API 미지원 버전입니다. Legacy purge로 진행합니다.");
                return 0;
            }

            var exclusions = new HashSet<ElementId>();
            object result = method.Invoke(doc, new object[] { exclusions });
            List<ElementId> ids = ToElementIdList(result)
                .Where(x => x != null && x != ElementId.InvalidElementId)
                .Where(x => context.KeepViewId == null || x.IntegerValue != context.KeepViewId.IntegerValue)
                .Where(x => context.KeepFilterId == null || x.IntegerValue != context.KeepFilterId.IntegerValue)
                .Distinct(new ElementIdComparer())
                .ToList();

            if (ids.Count == 0)
            {
                log?.Invoke("Native purge 후보 없음");
                return 0;
            }

            return DeleteIdsBatchFirst(doc, ids, "Native Purge", log, context.KeepViewId, context.KeepFilterId, context.ModelElementIdSet, 64, false, false, null);
        }

        private static List<ElementId> ToElementIdList(object result)
        {
            var list = new List<ElementId>();
            if (result == null) return list;

            var enumerable = result as System.Collections.IEnumerable;
            if (enumerable == null) return list;

            foreach (object item in enumerable)
            {
                if (item is ElementId id)
                {
                    list.Add(id);
                }
            }

            return list;
        }

        private static void LogPurgeScope(Document doc, PurgePassContext context, int pass, Action<string> log)
        {
            if (log == null) return;

            int textTypeCount = new FilteredElementCollector(doc).OfClass(typeof(TextNoteType)).GetElementCount();
            int filterCount = new FilteredElementCollector(doc).OfClass(typeof(FilterElement)).GetElementCount();
            int parameterElementCount = new FilteredElementCollector(doc).OfClass(typeof(ParameterElement)).GetElementCount();
            int materialCount = new FilteredElementCollector(doc).OfClass(typeof(Material)).GetElementCount();
            int linePatternCount = new FilteredElementCollector(doc).OfClass(typeof(LinePatternElement)).GetElementCount();
            int fillPatternCount = new FilteredElementCollector(doc).OfClass(typeof(FillPatternElement)).GetElementCount();
            int elementTypeCount = new FilteredElementCollector(doc).WhereElementIsElementType().GetElementCount();
            int unusedElementTypeCount = context.UnusedElementTypeCandidates != null ? context.UnusedElementTypeCandidates.Count : 0;
            int probeCount = context.ProbeCandidates != null ? context.ProbeCandidates.Count : 0;

            log($"Purge pass {pass} 후보 / ElementType {elementTypeCount}, UnusedElementType {unusedElementTypeCount}, TextType {textTypeCount}, Filter {filterCount}, ParameterElement {parameterElementCount}, Material {materialCount}, LinePattern {linePatternCount}, FillPattern {fillPatternCount}, ProbeCandidate {probeCount}");
        }

        private static int DeleteUnusedElementTypes(Document doc, PurgePassContext context, Action<string> log)
        {
            List<ElementId> candidates = context.UnusedElementTypeCandidates ?? new List<ElementId>();
            if (candidates.Count == 0)
            {
                log?.Invoke("미사용 ElementType 후보 없음");
                return 0;
            }

            return DeleteIdsBatchFirst(doc, candidates, "Delete Unused Element Types", log, context.KeepViewId, null, context.ModelElementIdSet, 128, true, false, null);
        }

        private static List<ElementId> CollectUnusedElementTypeCandidates(Document doc, ElementId keepViewId, HashSet<int> usedTypeIds)
        {
            usedTypeIds = usedTypeIds ?? BuildUsedTypeIdSet(doc);
            ElementId keepViewTypeId = ElementId.InvalidElementId;
            View keepView = keepViewId != null && keepViewId != ElementId.InvalidElementId ? doc.GetElement(keepViewId) as View : null;
            if (keepView != null && keepView.IsValidObject)
            {
                keepViewTypeId = keepView.GetTypeId();
            }

            return new FilteredElementCollector(doc)
                .WhereElementIsElementType()
                .Cast<ElementType>()
                .Where(x => x != null)
                .Where(x => keepViewTypeId == null || keepViewTypeId == ElementId.InvalidElementId || x.Id.IntegerValue != keepViewTypeId.IntegerValue)
                .Where(x => !usedTypeIds.Contains(x.Id.IntegerValue))
                .Where(IsPurgeableElementTypeCandidate)
                .Select(x => x.Id)
                .Distinct(new ElementIdComparer())
                .ToList();
        }

        private static HashSet<int> BuildUsedTypeIdSet(Document doc)
        {
            var usedTypeIds = new HashSet<int>();

            foreach (Element element in new FilteredElementCollector(doc).WhereElementIsNotElementType())
            {
                if (element == null) continue;

                try
                {
                    ElementId typeId = element.GetTypeId();
                    if (typeId != null && typeId != ElementId.InvalidElementId)
                    {
                        usedTypeIds.Add(typeId.IntegerValue);
                    }
                }
                catch
                {
                }
            }

            return usedTypeIds;
        }

        private static HashSet<int> BuildUsedFilterIdSet(Document doc)
        {
            var usedFilterIds = new HashSet<int>();
            foreach (ElementId viewId in new FilteredElementCollector(doc).OfClass(typeof(View)).ToElementIds())
            {
                View view = doc.GetElement(viewId) as View;
                if (view == null || !view.IsValidObject) continue;

                try
                {
                    foreach (ElementId filterId in view.GetFilters())
                    {
                        if (filterId != null && filterId != ElementId.InvalidElementId)
                        {
                            usedFilterIds.Add(filterId.IntegerValue);
                        }
                    }
                }
                catch
                {
                }
            }
            return usedFilterIds;
        }

        private static bool IsPurgeableElementTypeCandidate(ElementType type)
        {
            if (type == null) return false;
            if (type is ViewFamilyType) return false;
            if (type is TextNoteType) return false;
            if (type is RevitLinkType) return false;
            if (type is CADLinkType) return false;

            string typeName = type.GetType().Name;
            if (string.Equals(typeName, "ImageType", StringComparison.OrdinalIgnoreCase)) return false;
            if (string.Equals(typeName, "LevelType", StringComparison.OrdinalIgnoreCase)) return false;
            if (string.Equals(typeName, "GridType", StringComparison.OrdinalIgnoreCase)) return false;

            return true;
        }

        private static int DeleteUnusedTextNoteTypes(Document doc, PurgePassContext context, Action<string> log)
        {
            HashSet<int> usedTypeIds = context.UsedTypeIds ?? new HashSet<int>();
            List<ElementId> deletable = new FilteredElementCollector(doc)
                .OfClass(typeof(TextNoteType))
                .Cast<TextNoteType>()
                .Where(x => x != null)
                .Where(x => !usedTypeIds.Contains(x.Id.IntegerValue))
                .Select(x => x.Id)
                .Distinct(new ElementIdComparer())
                .ToList();

            if (deletable.Count == 0)
            {
                log?.Invoke("미사용 문자(TextNoteType) 후보 없음");
                return 0;
            }

            return DeleteIdsBatchFirst(doc, deletable, "Delete Unused Text Types", log, null, null, context.ModelElementIdSet, 64, true, false, null);
        }

        private static int DeleteUnusedFilters(Document doc, PurgePassContext context, Action<string> log)
        {
            HashSet<int> usedFilterIds = context.UsedFilterIds ?? new HashSet<int>();
            ElementId keepFilterId = context.KeepFilterId;

            List<ElementId> unusedFilterIds = new FilteredElementCollector(doc)
                .OfClass(typeof(FilterElement))
                .Cast<FilterElement>()
                .Where(x => x != null)
                .Where(x => keepFilterId == null || x.Id.IntegerValue != keepFilterId.IntegerValue)
                .Where(x => !usedFilterIds.Contains(x.Id.IntegerValue))
                .Select(x => x.Id)
                .Distinct(new ElementIdComparer())
                .ToList();

            if (unusedFilterIds.Count == 0)
            {
                log?.Invoke("미사용 뷰 필터 후보 없음");
                return 0;
            }

            return DeleteIdsBatchFirst(doc, unusedFilterIds, "Delete Unused View Filters", log, null, keepFilterId, context.ModelElementIdSet, 64, true, false, null);
        }

        private static int DeleteProbeSafeCandidates(Document doc, PurgePassContext context, Action<string> log)
        {
            List<ElementId> candidates = context.ProbeCandidates ?? new List<ElementId>();
            if (candidates.Count == 0)
            {
                log?.Invoke("Probe purge 후보 없음");
                return 0;
            }

            return DeleteIdsIndividually(doc, candidates, "Delete Probe Safe Candidates", log, context.KeepViewId, context.KeepFilterId, context.SafeDeletedIdSet, true, context.ModelElementIdSet);
        }

        private static List<ElementId> CollectGenericPurgeCandidates(Document doc, ElementId keepFilterId)
        {
            var result = new List<ElementId>();

            foreach (Element material in new FilteredElementCollector(doc).OfClass(typeof(Material)).ToElements())
            {
                if (material != null) result.Add(material.Id);
            }

            foreach (Element linePattern in new FilteredElementCollector(doc).OfClass(typeof(LinePatternElement)).ToElements())
            {
                if (linePattern != null) result.Add(linePattern.Id);
            }

            foreach (Element fillPattern in new FilteredElementCollector(doc).OfClass(typeof(FillPatternElement)).ToElements())
            {
                if (fillPattern != null) result.Add(fillPattern.Id);
            }

            foreach (ParameterElement parameterElement in new FilteredElementCollector(doc).OfClass(typeof(ParameterElement)).Cast<ParameterElement>())
            {
                if (parameterElement != null) result.Add(parameterElement.Id);
            }

            Type appearanceType = typeof(Document).Assembly.GetType("Autodesk.Revit.DB.AppearanceAssetElement", false);
            if (appearanceType != null)
            {
                foreach (Element appearance in new FilteredElementCollector(doc).OfClass(appearanceType).ToElements())
                {
                    if (appearance != null) result.Add(appearance.Id);
                }
            }

            return result.Distinct(new ElementIdComparer()).ToList();
        }

        private static HashSet<int> BuildSafeDeletedIdSet(Document doc, ElementId keepFilterId)
        {
            var set = new HashSet<int>();

            foreach (ElementType type in new FilteredElementCollector(doc).WhereElementIsElementType().Cast<ElementType>())
            {
                if (type != null) set.Add(type.Id.IntegerValue);
            }

            foreach (Element material in new FilteredElementCollector(doc).OfClass(typeof(Material)).ToElements())
            {
                if (material != null) set.Add(material.Id.IntegerValue);
            }

            foreach (Element linePattern in new FilteredElementCollector(doc).OfClass(typeof(LinePatternElement)).ToElements())
            {
                if (linePattern != null) set.Add(linePattern.Id.IntegerValue);
            }

            foreach (Element fillPattern in new FilteredElementCollector(doc).OfClass(typeof(FillPatternElement)).ToElements())
            {
                if (fillPattern != null) set.Add(fillPattern.Id.IntegerValue);
            }

            foreach (FilterElement filter in new FilteredElementCollector(doc).OfClass(typeof(FilterElement)).Cast<FilterElement>())
            {
                if (filter == null) continue;
                if (keepFilterId != null && keepFilterId != ElementId.InvalidElementId && filter.Id.IntegerValue == keepFilterId.IntegerValue) continue;
                set.Add(filter.Id.IntegerValue);
            }

            foreach (ParameterElement parameterElement in new FilteredElementCollector(doc).OfClass(typeof(ParameterElement)).Cast<ParameterElement>())
            {
                if (parameterElement != null) set.Add(parameterElement.Id.IntegerValue);
            }

            Type appearanceType = typeof(Document).Assembly.GetType("Autodesk.Revit.DB.AppearanceAssetElement", false);
            if (appearanceType != null)
            {
                foreach (Element element in new FilteredElementCollector(doc).OfClass(appearanceType).ToElements())
                {
                    if (element != null) set.Add(element.Id.IntegerValue);
                }
            }

            return set;
        }

        private static bool IsSafeDeletedSet(ElementId candidateId, ICollection<ElementId> deletedIds, ElementId keepViewId, ElementId keepFilterId, ISet<int> safeDeletedIdSet)
        {
            if (candidateId == null || candidateId == ElementId.InvalidElementId) return false;
            if (deletedIds == null || deletedIds.Count == 0) return false;

            foreach (ElementId deletedId in deletedIds)
            {
                if (deletedId == null || deletedId == ElementId.InvalidElementId) return false;
                if (keepViewId != null && keepViewId != ElementId.InvalidElementId && deletedId.IntegerValue == keepViewId.IntegerValue) return false;
                if (keepFilterId != null && keepFilterId != ElementId.InvalidElementId && deletedId.IntegerValue == keepFilterId.IntegerValue) return false;
                if (deletedId.IntegerValue == candidateId.IntegerValue) continue;
                if (safeDeletedIdSet != null && !safeDeletedIdSet.Contains(deletedId.IntegerValue)) return false;
            }

            return true;
        }

        private static int DeleteIdsBatchFirst(Document doc, IList<ElementId> ids, string transactionName, Action<string> log, ElementId keepViewId, ElementId keepFilterId, ISet<int> modelElementIdSet, int batchSize, bool fallbackToIndividual, bool probeFirst, ISet<int> safeDeletedIdSet)
        {
            if (ids == null || ids.Count == 0) return 0;

            List<ElementId> distinctIds = ids.Where(x => x != null && x != ElementId.InvalidElementId).Distinct(new ElementIdComparer()).ToList();
            int totalDeleted = 0;
            int failedBatches = 0;

            for (int i = 0; i < distinctIds.Count; i += Math.Max(1, batchSize))
            {
                List<ElementId> batch = distinctIds.Skip(i).Take(Math.Max(1, batchSize)).ToList();
                bool batchSucceeded = false;

                using (var tx = new Transaction(doc, transactionName + " Batch"))
                {
                    try
                    {
                        tx.Start();
                        AttachFailureSwallower(tx);

                        List<ElementId> liveIds = batch.Where(x => doc.GetElement(x) != null).ToList();
                        if (liveIds.Count == 0)
                        {
                            tx.RollBack();
                            batchSucceeded = true;
                        }
                        else
                        {
                            ICollection<ElementId> removed = doc.Delete(liveIds);
                            int modelRisk = CountModelDeletionRisk(removed, modelElementIdSet, keepViewId);
                            bool safeDeleteSet = !probeFirst || IsBatchSafeDeletedSet(liveIds, removed, keepViewId, keepFilterId, safeDeletedIdSet);
                            if (modelRisk > 0 || !safeDeleteSet)
                            {
                                tx.RollBack();
                            }
                            else
                            {
                                totalDeleted += removed != null ? removed.Count : 0;
                                tx.Commit();
                                batchSucceeded = true;
                            }
                        }
                    }
                    catch
                    {
                        try { tx.RollBack(); } catch { }
                    }
                }

                if (!batchSucceeded)
                {
                    failedBatches++;
                    if (fallbackToIndividual)
                    {
                        totalDeleted += DeleteIdsIndividually(doc, batch, transactionName + " Fallback", log, keepViewId, keepFilterId, safeDeletedIdSet, probeFirst, modelElementIdSet);
                    }
                }
            }

            if (totalDeleted > 0)
            {
                log?.Invoke($"{transactionName} → {totalDeleted}개 삭제 (배치 우선, 개별 fallback)");
            }
            else
            {
                log?.Invoke($"{transactionName} → 삭제 없음 (검토 {distinctIds.Count}, 실패 배치 {failedBatches})");
            }

            return totalDeleted;
        }

        private static bool IsBatchSafeDeletedSet(IList<ElementId> candidateIds, ICollection<ElementId> deletedIds, ElementId keepViewId, ElementId keepFilterId, ISet<int> safeDeletedIdSet)
        {
            if (candidateIds == null || candidateIds.Count == 0) return false;
            if (deletedIds == null || deletedIds.Count == 0) return false;

            HashSet<int> candidateSet = new HashSet<int>(candidateIds.Where(x => x != null && x != ElementId.InvalidElementId).Select(x => x.IntegerValue));
            foreach (ElementId deletedId in deletedIds)
            {
                if (deletedId == null || deletedId == ElementId.InvalidElementId) return false;
                if (keepViewId != null && keepViewId != ElementId.InvalidElementId && deletedId.IntegerValue == keepViewId.IntegerValue) return false;
                if (keepFilterId != null && keepFilterId != ElementId.InvalidElementId && deletedId.IntegerValue == keepFilterId.IntegerValue) return false;
                if (candidateSet.Contains(deletedId.IntegerValue)) continue;
                if (safeDeletedIdSet != null && !safeDeletedIdSet.Contains(deletedId.IntegerValue)) return false;
            }

            return true;
        }

        private static int DeleteIdsIndividually(Document doc, IList<ElementId> ids, string transactionName, Action<string> log, ElementId keepViewId, ElementId keepFilterId, ISet<int> safeDeletedIdSet, bool probeFirst = false, ISet<int> modelElementIdSet = null)
        {
            if (ids == null || ids.Count == 0) return 0;

            int deleted = 0;
            int failed = 0;
            int modelRiskSkip = 0;
            var deletedByClass = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

            foreach (ElementId id in ids)
            {
                if (id == null || id == ElementId.InvalidElementId) continue;

                Element snapshotElement = doc.GetElement(id);
                if (snapshotElement == null) continue;

                int candidateIdInt = id.IntegerValue;
                string candidateClassName = GetPurgeLogClassName(snapshotElement);
                bool shouldDelete = !probeFirst;
                int modelDeletionRiskCount = 0;

                if (probeFirst)
                {
                    using (var probeTx = new Transaction(doc, transactionName + " Probe"))
                    {
                        try
                        {
                            probeTx.Start();
                            AttachFailureSwallower(probeTx);

                            ICollection<ElementId> probeDeleted = doc.Delete(id);
                            modelDeletionRiskCount = CountModelDeletionRisk(probeDeleted, modelElementIdSet, keepViewId);
                            shouldDelete = modelDeletionRiskCount == 0 && IsSafeDeletedSet(id, probeDeleted, keepViewId, keepFilterId, safeDeletedIdSet);
                            probeTx.RollBack();
                        }
                        catch
                        {
                            try { probeTx.RollBack(); } catch { }
                            shouldDelete = false;
                        }
                    }
                }

                if (!shouldDelete)
                {
                    if (modelDeletionRiskCount > 0)
                    {
                        modelRiskSkip++;
                        log?.Invoke($"{transactionName} 모델 객체 삭제 위험으로 패스: {candidateClassName} {candidateIdInt} / 영향 모델 객체 {modelDeletionRiskCount}개");
                    }

                    failed++;
                    continue;
                }

                using (var deleteTx = new Transaction(doc, transactionName + " Delete"))
                {
                    try
                    {
                        deleteTx.Start();
                        AttachFailureSwallower(deleteTx);

                        if (doc.GetElement(id) == null)
                        {
                            deleteTx.RollBack();
                            continue;
                        }

                        ICollection<ElementId> removed = doc.Delete(id);
                        if (probeFirst)
                        {
                            modelDeletionRiskCount = CountModelDeletionRisk(removed, modelElementIdSet, keepViewId);
                        }

                        if ((probeFirst && !IsSafeDeletedSet(id, removed, keepViewId, keepFilterId, safeDeletedIdSet)) || modelDeletionRiskCount > 0)
                        {
                            if (modelDeletionRiskCount > 0)
                            {
                                modelRiskSkip++;
                                log?.Invoke($"{transactionName} 실제 삭제 단계 모델 객체 영향 감지로 롤백: {candidateClassName} {candidateIdInt} / 영향 모델 객체 {modelDeletionRiskCount}개");
                            }

                            deleteTx.RollBack();
                            failed++;
                            continue;
                        }

                        int removedCount = removed != null ? removed.Count : 0;
                        if (removedCount > 0)
                        {
                            deleted += removedCount;
                            if (!deletedByClass.ContainsKey(candidateClassName))
                            {
                                deletedByClass[candidateClassName] = 0;
                            }
                            deletedByClass[candidateClassName] += removedCount;
                        }

                        deleteTx.Commit();
                    }
                    catch
                    {
                        try { deleteTx.RollBack(); } catch { }
                        failed++;
                    }
                }
            }

            if (deleted > 0)
            {
                string summary = string.Join(", ", deletedByClass.OrderByDescending(x => x.Value).Select(x => x.Key + " " + x.Value));
                log?.Invoke($"{transactionName} → {deleted}개 삭제 / {summary}");
            }
            else
            {
                log?.Invoke($"{transactionName} → 삭제 없음 (검토 {ids.Count}, 스킵/실패 {failed}, 모델영향패스 {modelRiskSkip})");
            }

            return deleted;
        }

        private static HashSet<int> BuildModelElementIdSet(Document doc)
        {
            var set = new HashSet<int>();

            foreach (Element element in new FilteredElementCollector(doc).WhereElementIsNotElementType())
            {
                if (element == null) continue;
                if (element.ViewSpecific) continue;

                Category category = element.Category;
                if (category == null) continue;

                try
                {
                    if (category.CategoryType == CategoryType.Model)
                    {
                        set.Add(element.Id.IntegerValue);
                    }
                }
                catch
                {
                }
            }

            return set;
        }

        private static int CountModelDeletionRisk(ICollection<ElementId> deletedIds, ISet<int> modelElementIdSet, ElementId keepViewId)
        {
            if (deletedIds == null || deletedIds.Count == 0) return 0;
            if (modelElementIdSet == null || modelElementIdSet.Count == 0) return 0;

            int count = 0;
            foreach (ElementId deletedId in deletedIds)
            {
                if (deletedId == null || deletedId == ElementId.InvalidElementId) continue;
                if (keepViewId != null && keepViewId != ElementId.InvalidElementId && deletedId.IntegerValue == keepViewId.IntegerValue) continue;
                if (modelElementIdSet.Contains(deletedId.IntegerValue))
                {
                    count++;
                }
            }

            return count;
        }

        private static string GetPurgeLogClassName(Element element)
        {
            if (element == null) return "Unknown";
            Type type = element.GetType();
            return type != null ? type.Name : "Unknown";
        }

        private static void AttachFailureSwallower(Transaction tx)
        {
            FailureHandlingOptions options = tx.GetFailureHandlingOptions();
            options.SetFailuresPreprocessor(new TransactionFailureSwallower());
            tx.SetFailureHandlingOptions(options);
        }

        private sealed class PurgePassContext
        {
            public ElementId KeepViewId { get; set; }
            public ElementId KeepFilterId { get; set; }
            public HashSet<int> UsedTypeIds { get; set; }
            public HashSet<int> UsedFilterIds { get; set; }
            public HashSet<int> ModelElementIdSet { get; set; }
            public HashSet<int> SafeDeletedIdSet { get; set; }
            public List<ElementId> UnusedElementTypeCandidates { get; set; }
            public List<ElementId> ProbeCandidates { get; set; }
        }

        private sealed class ElementIdComparer : IEqualityComparer<ElementId>
        {
            public bool Equals(ElementId x, ElementId y)
            {
                if (ReferenceEquals(x, y)) return true;
                if (ReferenceEquals(x, null) || ReferenceEquals(y, null)) return false;
                return x.IntegerValue == y.IntegerValue;
            }

            public int GetHashCode(ElementId obj)
            {
                return obj?.IntegerValue ?? 0;
            }
        }
    }
}
