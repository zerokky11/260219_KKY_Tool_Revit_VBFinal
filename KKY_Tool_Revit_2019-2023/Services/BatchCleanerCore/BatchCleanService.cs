using System;
using System.Collections.Generic;
using System.IO;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Text;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KKY_Tool_Revit.Models;

namespace KKY_Tool_Revit.Services
{
    public static class BatchCleanService
    {
        public static BatchPrepareSession Prepare(UIApplication uiapp, BatchCleanSettings settings, Action<string> log)
        {
            if (uiapp == null) throw new ArgumentNullException(nameof(uiapp));
            if (settings == null) throw new ArgumentNullException(nameof(settings));

            var session = new BatchPrepareSession
            {
                OutputFolder = settings.OutputFolder
            };
            var designOptionAuditRows = new List<DesignOptionAuditRow>();

            if (!string.IsNullOrWhiteSpace(settings.OutputFolder))
            {
                Directory.CreateDirectory(settings.OutputFolder);
            }

            foreach (string filePath in settings.FilePaths.Where(File.Exists))
            {
                var result = new BatchCleanResult { SourcePath = filePath };
                session.Results.Add(result);

                try
                {
                    log?.Invoke("----------------------------------------");
                    log?.Invoke("?뺣━ ?쒖옉: " + filePath);

                    PreparedDocumentEntry prepared = PrepareSingle(uiapp, filePath, settings, designOptionAuditRows, log);
                    session.PreparedDocuments.Add(prepared);
                    session.CleanCountComparisons.Add(BuildCleanCountComparison(prepared, true, null));
                    result.OutputPath = prepared.OutputPath;
                    result.Success = true;
                    result.Message = "정리 완료 / 저장 완료";

                    log?.Invoke("정리 완료(저장 경로): " + prepared.OutputPath);
                }
                catch (Exception ex)
                {
                    session.CleanCountComparisons.Add(BuildCleanCountComparison(filePath, null, false, ex.Message));
                    result.Success = false;
                    result.Message = ex.Message;
                    log?.Invoke("?ㅽ뙣: " + ex.Message);
                }
            }

            try
            {
                string auditPath = WriteDesignOptionAuditCsv(settings.OutputFolder, designOptionAuditRows);
                session.DesignOptionAuditCsvPath = auditPath;
                if (!string.IsNullOrWhiteSpace(auditPath))
                {
                    log?.Invoke("Design Option CSV ??? " + auditPath);
                }
            }
            catch (Exception ex)
            {
                log?.Invoke("Design Option CSV ????ㅽ뙣: " + ex.Message);
            }

            return session;
        }

        public static BatchPrepareSession CleanAndSave(UIApplication uiapp, BatchCleanSettings settings, Action<string> log)
        {
            if (uiapp == null) throw new ArgumentNullException(nameof(uiapp));
            if (settings == null) throw new ArgumentNullException(nameof(settings));

            var session = new BatchPrepareSession
            {
                OutputFolder = settings.OutputFolder
            };
            var designOptionAuditRows = new List<DesignOptionAuditRow>();

            if (!string.IsNullOrWhiteSpace(settings.OutputFolder))
            {
                Directory.CreateDirectory(settings.OutputFolder);
            }

            foreach (string filePath in settings.FilePaths.Where(File.Exists))
            {
                var result = new BatchCleanResult { SourcePath = filePath };
                session.Results.Add(result);

                PreparedDocumentEntry prepared = null;
                try
                {
                    log?.Invoke("----------------------------------------");
                    log?.Invoke("?뺣━ ?쒖옉: " + filePath);

                    prepared = PrepareSingle(uiapp, filePath, settings, designOptionAuditRows, log);
                    result.OutputPath = prepared.OutputPath;

                    log?.Invoke("[STEP] Groups 釉뚮씪?곗? 媛깆떊 ?좊룄 ?쒖옉");
                    try
                    {
                        TouchGroupsBrowser(prepared.Document, log);
                    }
                    catch (Exception ex)
                    {
                        log?.Invoke("Groups 媛깆떊 ?좊룄 ?ㅽ뙣(臾댁떆): " + ex.Message);
                    }

                    log?.Invoke("[STEP] ????쒖옉");
                    SaveCleanDocument(prepared.Document, prepared.TargetViewId, prepared.OutputPath, log);
                    session.CleanedOutputPaths.Add(prepared.OutputPath);
                    session.CleanCountComparisons.Add(BuildCleanCountComparison(prepared, true, null));

                    result.Success = true;
                    result.Message = "?뺣━ 諛?????꾨즺";
                    log?.Invoke("?뺣━ 諛?????꾨즺: " + prepared.OutputPath);
                }
                catch (Exception ex)
                {
                    session.CleanCountComparisons.Add(BuildCleanCountComparison(filePath, prepared, false, ex.Message));
                    result.Success = false;
                    result.Message = ex.Message;
                    log?.Invoke("?ㅽ뙣: " + ex.Message);
                }
                finally
                {
                    try
                    {
                        if (prepared != null && prepared.Document != null && prepared.Document.IsValidObject)
                        {
                            prepared.Document.Close(false);
                        }
                    }
                    catch
                    {
                    }
                }
            }

            try
            {
                string auditPath = WriteDesignOptionAuditCsv(settings.OutputFolder, designOptionAuditRows);
                session.DesignOptionAuditCsvPath = auditPath;
                if (!string.IsNullOrWhiteSpace(auditPath))
                {
                    log?.Invoke("Design Option CSV ??? " + auditPath);
                }
            }
            catch (Exception ex)
            {
                log?.Invoke("Design Option CSV ????ㅽ뙣: " + ex.Message);
            }

            return session;
        }

        public static List<BatchCleanResult> SavePrepared(BatchPrepareSession session, Action<string> log)
        {
            if (session == null) throw new ArgumentNullException(nameof(session));

            var results = new List<BatchCleanResult>();
            foreach (PreparedDocumentEntry prepared in session.PreparedDocuments)
            {
                var result = new BatchCleanResult
                {
                    SourcePath = prepared.SourcePath,
                    OutputPath = prepared.OutputPath
                };
                results.Add(result);

                Document doc = prepared.Document;
                try
                {
                    if (doc == null || !doc.IsValidObject)
                    {
                        throw new InvalidOperationException("??????臾몄꽌媛 ?좏슚?섏? ?딆뒿?덈떎.");
                    }

                    log?.Invoke("----------------------------------------");
                    log?.Invoke("????쒖옉: " + prepared.SourcePath);
                    log?.Invoke("[STEP] Groups 釉뚮씪?곗? 媛깆떊 ?좊룄 ?쒖옉");
                    try
                    {
                        TouchGroupsBrowser(doc, log);
                    }
                    catch (Exception ex)
                    {
                        log?.Invoke("Groups 媛깆떊 ?좊룄 ?ㅽ뙣(臾댁떆): " + ex.Message);
                    }

                    log?.Invoke("[STEP] ????쒖옉");
                    SaveCleanDocument(doc, prepared.TargetViewId, prepared.OutputPath, log);
                    result.Success = true;
                    result.Message = "????꾨즺";
                }
                catch (Exception ex)
                {
                    result.Success = false;
                    result.Message = ex.Message;
                    log?.Invoke("????ㅽ뙣: " + ex.Message);
                }
                finally
                {
                    if (doc != null && doc.IsValidObject)
                    {
                        try
                        {
                            doc.Close(false);
                        }
                        catch
                        {
                            // ignore
                        }
                    }
                }
            }

            session.PreparedDocuments.Clear();
            return results;
        }

        public static void ClosePreparedWithoutSaving(BatchPrepareSession session, Action<string> log)
        {
            if (session == null) return;

            foreach (PreparedDocumentEntry prepared in session.PreparedDocuments)
            {
                try
                {
                    if (prepared.Document != null && prepared.Document.IsValidObject)
                    {
                        prepared.Document.Close(false);
                        log?.Invoke("????놁씠 ?レ쓬: " + prepared.SourcePath);
                    }
                }
                catch
                {
                    // ignore
                }
            }

            session.PreparedDocuments.Clear();
        }

        public static List<BatchCleanResult> Run(UIApplication uiapp, BatchCleanSettings settings, Action<string> log)
        {
            BatchPrepareSession session = CleanAndSave(uiapp, settings, log);
            return session.Results;
        }

        private static PreparedDocumentEntry PrepareSingle(UIApplication uiapp, string filePath, BatchCleanSettings settings, List<DesignOptionAuditRow> designOptionAuditRows, Action<string> log)
        {
            Document doc = null;
            try
            {
                doc = OpenDetachedOrNormal(uiapp, filePath, log);
                if (doc == null) throw new InvalidOperationException("臾몄꽌瑜??댁? 紐삵뻽?듬땲??");

                int beforeObjectCount = ModelParameterExtractionService.CountExtractableElements(doc);
                IDictionary<string, int> beforeObjectBreakdown = ModelParameterExtractionService.GetExtractableElementSignatureCounts(doc);
                CollectDesignOptionAuditRows(doc, filePath, designOptionAuditRows, log);

                log?.Invoke("[STEP] Manage Links ?뺣━ ?쒖옉");
                DeleteManageLinks(doc, log);
                log?.Invoke("[STEP] Group/Assembly ?댁젣 ?쒖옉");
                UngroupAndDisassemble(doc, log);
                log?.Invoke("[STEP] 湲곗〈 3D 酉??숇챸 酉???젣 ?쒖옉");
                DeleteExisting3DViewsAndConflictingNamedViews(doc, settings.Target3DViewName, log);

                log?.Invoke("[STEP] ???3D 酉??앹꽦 ?쒖옉");
                ElementId targetViewId = CreateTarget3DView(doc, settings.Target3DViewName, log);
                log?.Invoke("[STEP] ???3D 酉??ㅼ젙 ?곸슜 ?쒖옉");
                ConfigureTarget3DView(doc, targetViewId, settings, log);

                ElementId keptFilterId = ElementId.InvalidElementId;
                if (settings.UseFilter && settings.FilterProfile != null && settings.FilterProfile.IsConfigured())
                {
                    keptFilterId = CreateOrApplyFilter(doc, targetViewId, settings, log);
                }

                log?.Invoke("[STEP] ???酉????섎㉧吏 酉??쒗뵆由???젣 ?쒖옉");
                DeleteAllOtherViewsAndTemplates(doc, targetViewId, log);
                log?.Invoke("[STEP] Starting View ?ㅼ젙 ?쒖옉");
                SetStartingView(doc, targetViewId, log);
                log?.Invoke("[STEP] 誘몄궗??酉??꾪꽣 ??젣 ?쒖옉");
                DeleteUnusedViewFilters(doc, keptFilterId, log);

                if (settings.ElementParameterUpdate != null && settings.ElementParameterUpdate.IsConfigured())
                {
                    log?.Invoke("[STEP] 媛앹껜 ?뚮씪誘명꽣 ?쇨큵 ?낅젰 ?쒖옉");
                    ApplyElementParameterUpdate(doc, settings.ElementParameterUpdate, log);
                }
                else
                {
                    log?.Invoke("[STEP] 媛앹껜 ?뚮씪誘명꽣 ?쇨큵 ?낅젰 嫄대꼫?");
                }

                log?.Invoke("[STEP] Legacy Purge ?쒓굅??- ?뺣━ ?④퀎?먯꽌??Purge瑜??ㅽ뻾?섏? ?딆쓬");

                int afterObjectCount = ModelParameterExtractionService.CountExtractableElements(doc);
                IDictionary<string, int> afterObjectBreakdown = ModelParameterExtractionService.GetExtractableElementSignatureCounts(doc);
                string removedSummary = string.Empty;
                if (afterObjectCount < beforeObjectCount)
                {
                    removedSummary = ModelParameterExtractionService.BuildReductionSummary(beforeObjectBreakdown, afterObjectBreakdown);
                    if (!string.IsNullOrWhiteSpace(removedSummary))
                    {
                        log?.Invoke("[STEP] 媛먯냼 媛앹껜 ?붿빟 / " + removedSummary);
                    }
                }
                log?.Invoke("[STEP] 媛앹껜??鍮꾧탳 / ?뺣━ ??" + beforeObjectCount + ", ?뺣━ ??" + afterObjectCount);

                return new PreparedDocumentEntry
                {
                    SourcePath = filePath,
                    OutputPath = BuildOutputPath(filePath, settings.OutputFolder),
                    Document = doc,
                    TargetViewId = targetViewId,
                    KeptFilterId = keptFilterId,
                    BeforeObjectCount = beforeObjectCount,
                    AfterObjectCount = afterObjectCount,
                    RemovedSummary = removedSummary
                };
            }
            catch
            {
                if (doc != null && doc.IsValidObject)
                {
                    try
                    {
                        doc.Close(false);
                    }
                    catch
                    {
                        // ignore
                    }
                }

                throw;
            }
        }

        private static Document OpenDetachedOrNormal(UIApplication uiapp, string filePath, Action<string> log)
        {
            ModelPath modelPath = ModelPathUtils.ConvertUserVisiblePathToModelPath(filePath);

            try
            {
                var openOptions = new OpenOptions
                {
                    DetachFromCentralOption = DetachFromCentralOption.DetachAndDiscardWorksets
                };
                openOptions.SetOpenWorksetsConfiguration(new WorksetConfiguration(WorksetConfigurationOption.CloseAllWorksets));

                log?.Invoke("Detach + 紐⑤뱺 ?뚰겕???リ린 + workset discard ?닿린 ?쒕룄");
                return uiapp.Application.OpenDocumentFile(modelPath, openOptions);
            }
            catch
            {
                var openOptions = new OpenOptions
                {
                    DetachFromCentralOption = DetachFromCentralOption.DoNotDetach
                };
                openOptions.SetOpenWorksetsConfiguration(new WorksetConfiguration(WorksetConfigurationOption.CloseAllWorksets));

                log?.Invoke("모든 워크셋 닫기 일반 열기로 재시도");
                return uiapp.Application.OpenDocumentFile(modelPath, openOptions);
            }
        }

        private static void DeleteManageLinks(Document doc, Action<string> log)
        {
            using (var tx = new Transaction(doc, "Delete Manage Links"))
            {
                tx.Start();
                AttachFailureSwallower(tx);

                var ids = new List<ElementId>();
                ids.AddRange(new FilteredElementCollector(doc).OfClass(typeof(RevitLinkInstance)).ToElementIds());
                ids.AddRange(new FilteredElementCollector(doc).OfClass(typeof(ImportInstance)).ToElementIds());
                ids.AddRange(new FilteredElementCollector(doc).OfClass(typeof(RevitLinkType)).ToElementIds());
                ids.AddRange(new FilteredElementCollector(doc).OfClass(typeof(CADLinkType)).ToElementIds());

                ids.AddRange(GetElementIdsByTypeName(doc, "Autodesk.Revit.DB.ImageInstance"));
                ids.AddRange(GetElementIdsByTypeName(doc, "Autodesk.Revit.DB.PointClouds.PointCloudInstance"));
                ids.AddRange(GetElementIdsByTypeName(doc, "Autodesk.Revit.DB.ImageType"));
                ids.AddRange(GetElementIdsByTypeName(doc, "Autodesk.Revit.DB.PointClouds.PointCloudType"));

                ids = ids
                    .Where(x => x != null && x != ElementId.InvalidElementId)
                    .Distinct(new ElementIdComparer())
                    .ToList();

                if (ids.Count > 0)
                {
                    int failedDeleteCount = 0;
                    int deletedCount = DeleteElementsIndividually(doc, ids, "ManageLinks", log, ref failedDeleteCount);
                    if (failedDeleteCount > 0 && deletedCount > 0)
                    {
                        int retryFailedDeleteCount = 0;
                        int retryDeletedCount = DeleteElementsIndividually(doc, ids, "ManageLinksRetry", log, ref retryFailedDeleteCount);
                        deletedCount += retryDeletedCount;
                        failedDeleteCount = retryFailedDeleteCount;
                    }
                    log?.Invoke("Manage Links/Imports ??젣 寃곌낵 / ???" + ids.Count + ", ?깃났 " + deletedCount + ", ?ㅽ뙣 " + failedDeleteCount);
                }

                tx.Commit();
                log?.Invoke("Manage Links/Imports ?④퀎 ?꾨즺");
                log?.Invoke("Manage Links/Imports ??젣 ??? " + ids.Count);
            }
        }

        private static int DeleteElementsIndividually(Document doc, IEnumerable<ElementId> ids, string label, Action<string> log, ref int failedCount)
        {
            List<ElementId> targets = (ids ?? Enumerable.Empty<ElementId>())
                .Where(x => x != null && x != ElementId.InvalidElementId)
                .Distinct(new ElementIdComparer())
                .ToList();

            if (targets.Count == 0)
            {
                return 0;
            }

            int deletedCount = 0;
            int localFailedCount = 0;
            int detailedFailureLogCount = 0;

            foreach (ElementId id in targets)
            {
                try
                {
                    if (doc.GetElement(id) == null)
                    {
                        continue;
                    }

                    doc.Delete(id);
                    deletedCount++;
                }
                catch (Exception ex)
                {
                    localFailedCount++;
                    if (detailedFailureLogCount < 3)
                    {
                        log?.Invoke(label + " ??젣 ?ㅽ뙣: Id " + id.IntegerValue + " / " + ex.Message);
                        detailedFailureLogCount++;
                    }
                }
            }

            failedCount += localFailedCount;
            log?.Invoke(label + " ??젣 ?쒕룄: ???" + targets.Count + ", ?깃났 " + deletedCount + ", ?ㅽ뙣 " + localFailedCount);
            return deletedCount;
        }

        private static List<ElementId> GetElementIdsByTypeName(Document doc, string fullTypeName)
        {
            var list = new List<ElementId>();
            Type type = typeof(Document).Assembly.GetType(fullTypeName, false);
            if (type == null) return list;
            list.AddRange(new FilteredElementCollector(doc).OfClass(type).ToElementIds());
            return list;
        }

        private static void UngroupAndDisassemble(Document doc, Action<string> log)
        {
            using (var tx = new Transaction(doc, "Ungroup and Disassemble"))
            {
                tx.Start();
                AttachFailureSwallower(tx);

                List<ElementId> initialGroupIds = new FilteredElementCollector(doc)
                    .OfClass(typeof(Group))
                    .ToElementIds()
                    .Distinct(new ElementIdComparer())
                    .ToList();

                List<ElementId> initialGroupTypeIds = new FilteredElementCollector(doc)
                    .OfClass(typeof(GroupType))
                    .ToElementIds()
                    .Distinct(new ElementIdComparer())
                    .ToList();

                List<ElementId> initialAssemblyIds = new FilteredElementCollector(doc)
                    .OfClass(typeof(AssemblyInstance))
                    .ToElementIds()
                    .Distinct(new ElementIdComparer())
                    .ToList();

                List<ElementId> initialAssemblyTypeIds = GetElementIdsByTypeName(doc, "Autodesk.Revit.DB.AssemblyType")
                    .Distinct(new ElementIdComparer())
                    .ToList();

                int groupUngroupCount = 0;
                int groupDeleteCount = 0;
                foreach (ElementId groupId in initialGroupIds)
                {
                    Group group = doc.GetElement(groupId) as Group;
                    if (group == null || !group.IsValidObject)
                    {
                        continue;
                    }

                    try
                    {
                        group.UngroupMembers();
                        groupUngroupCount++;
                    }
                    catch
                    {
                        try
                        {
                            if (doc.GetElement(groupId) != null)
                            {
                                doc.Delete(groupId);
                                groupDeleteCount++;
                            }
                        }
                        catch
                        {
                            // ignore
                        }
                    }
                }

                int assemblyDisassembleCount = 0;
                int assemblyDeleteCount = 0;
                foreach (ElementId assemblyId in initialAssemblyIds)
                {
                    AssemblyInstance assembly = doc.GetElement(assemblyId) as AssemblyInstance;
                    if (assembly == null || !assembly.IsValidObject)
                    {
                        continue;
                    }

                    try
                    {
                        assembly.Disassemble();
                        assemblyDisassembleCount++;
                    }
                    catch
                    {
                        try
                        {
                            if (doc.GetElement(assemblyId) != null)
                            {
                                doc.Delete(assemblyId);
                                assemblyDeleteCount++;
                            }
                        }
                        catch
                        {
                            // ignore
                        }
                    }
                }

                int leftoverGroupDeleteCount = 0;
                List<ElementId> leftoverGroupIds = new FilteredElementCollector(doc)
                    .OfClass(typeof(Group))
                    .ToElementIds()
                    .Distinct(new ElementIdComparer())
                    .ToList();

                foreach (ElementId leftoverGroupId in leftoverGroupIds)
                {
                    try
                    {
                        if (doc.GetElement(leftoverGroupId) != null)
                        {
                            doc.Delete(leftoverGroupId);
                            leftoverGroupDeleteCount++;
                        }
                    }
                    catch
                    {
                        // ignore
                    }
                }

                int leftoverAssemblyDeleteCount = 0;
                List<ElementId> leftoverAssemblyIds = new FilteredElementCollector(doc)
                    .OfClass(typeof(AssemblyInstance))
                    .ToElementIds()
                    .Distinct(new ElementIdComparer())
                    .ToList();

                foreach (ElementId leftoverAssemblyId in leftoverAssemblyIds)
                {
                    try
                    {
                        if (doc.GetElement(leftoverAssemblyId) != null)
                        {
                            doc.Delete(leftoverAssemblyId);
                            leftoverAssemblyDeleteCount++;
                        }
                    }
                    catch
                    {
                        // ignore
                    }
                }

                int groupTypeDeleteCount = 0;
                List<ElementId> allGroupTypeIds = initialGroupTypeIds
                    .Concat(new FilteredElementCollector(doc).OfClass(typeof(GroupType)).ToElementIds())
                    .Where(x => x != null && x != ElementId.InvalidElementId)
                    .Distinct(new ElementIdComparer())
                    .ToList();

                foreach (ElementId groupTypeId in allGroupTypeIds)
                {
                    try
                    {
                        if (doc.GetElement(groupTypeId) != null)
                        {
                            doc.Delete(groupTypeId);
                            groupTypeDeleteCount++;
                        }
                    }
                    catch
                    {
                        // ignore
                    }
                }

                int assemblyTypeDeleteCount = 0;
                List<ElementId> allAssemblyTypeIds = initialAssemblyTypeIds
                    .Concat(GetElementIdsByTypeName(doc, "Autodesk.Revit.DB.AssemblyType"))
                    .Where(x => x != null && x != ElementId.InvalidElementId)
                    .Distinct(new ElementIdComparer())
                    .ToList();

                foreach (ElementId assemblyTypeId in allAssemblyTypeIds)
                {
                    try
                    {
                        if (doc.GetElement(assemblyTypeId) != null)
                        {
                            doc.Delete(assemblyTypeId);
                            assemblyTypeDeleteCount++;
                        }
                    }
                    catch
                    {
                        // ignore
                    }
                }

                tx.Commit();
                log?.Invoke($"洹몃９ ?댁젣: {groupUngroupCount}, 洹몃９ ??젣: {groupDeleteCount + leftoverGroupDeleteCount}, 洹몃９ ?????젣: {groupTypeDeleteCount}, ?댁뀍釉붾━ ?댁젣: {assemblyDisassembleCount}, ?댁뀍釉붾━ ??젣: {assemblyDeleteCount + leftoverAssemblyDeleteCount}, ?댁뀍釉붾━ ?????젣: {assemblyTypeDeleteCount}");
            }
        }

        private static void DeleteExisting3DViewsAndConflictingNamedViews(Document doc, string requestedViewName, Action<string> log)
        {
            string exactName = GetRequestedViewName(requestedViewName);
            List<ElementId> candidateIds = new FilteredElementCollector(doc)
                .OfClass(typeof(View))
                .Cast<View>()
                .Where(x => x != null && ShouldDeleteBeforeCreatingTargetView(x, exactName))
                .Select(x => x.Id)
                .Distinct(new ElementIdComparer())
                .ToList();

            int deletedCount = 0;
            int failedCount = 0;

            using (var tx = new Transaction(doc, "Delete Existing 3D Views and Name Conflicts"))
            {
                tx.Start();
                AttachFailureSwallower(tx);

                foreach (ElementId viewId in candidateIds)
                {
                    if (viewId == null || viewId == ElementId.InvalidElementId)
                    {
                        continue;
                    }

                    try
                    {
                        if (doc.GetElement(viewId) == null)
                        {
                            continue;
                        }

                        doc.Delete(viewId);
                        deletedCount++;
                    }
                    catch
                    {
                        failedCount++;
                    }
                }

                tx.Commit();
            }

            if (ViewNameExists(doc, exactName))
            {
                throw new InvalidOperationException("?ъ슜??吏??3D 酉??대쫫??湲곗〈 酉곗? 異⑸룎?섏뿬 ?뺥솗???대쫫?쇰줈 留뚮뱾 ???놁뒿?덈떎: " + exactName);
            }

            log?.Invoke($"湲곗〈 3D 酉??숇챸 酉??좎궘?? ???{candidateIds.Count}, ?깃났 {deletedCount}, ?ㅽ뙣 {failedCount}");
        }

        private static bool ShouldDeleteBeforeCreatingTargetView(View view, string exactName)
        {
            if (view == null) return false;

            if (view is View3D && !view.IsTemplate)
            {
                return true;
            }

            if (string.Equals(view.Name, exactName, StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            return false;
        }

        private static ElementId CreateTarget3DView(Document doc, string requestedViewName, Action<string> log)
        {
            string exactName = GetRequestedViewName(requestedViewName);
            ElementId createdViewId;

            using (var tx = new Transaction(doc, "Create Target 3D View"))
            {
                tx.Start();
                AttachFailureSwallower(tx);

                if (ViewNameExists(doc, exactName))
                {
                    throw new InvalidOperationException("???3D 酉??대쫫???꾩쭅 臾몄꽌???⑥븘 ?덉뒿?덈떎: " + exactName);
                }

                ViewFamilyType threeDType = new FilteredElementCollector(doc)
                    .OfClass(typeof(ViewFamilyType))
                    .Cast<ViewFamilyType>()
                    .FirstOrDefault(x => x.ViewFamily == ViewFamily.ThreeDimensional);

                if (threeDType == null)
                {
                    throw new InvalidOperationException("3D ViewFamilyType??李얠? 紐삵뻽?듬땲??");
                }

                View3D createdView = View3D.CreateIsometric(doc, threeDType.Id);
                createdView.Name = exactName;
                createdView.ViewTemplateId = ElementId.InvalidElementId;
                createdViewId = createdView.Id;

                tx.Commit();
            }

            log?.Invoke("???3D 酉??앹꽦: " + exactName);
            return createdViewId;
        }

        private static void ConfigureTarget3DView(Document doc, ElementId targetViewId, BatchCleanSettings settings, Action<string> log)
        {
            using (var tx = new Transaction(doc, "Configure Target 3D View"))
            {
                tx.Start();
                AttachFailureSwallower(tx);

                View3D createdView = GetRequiredTargetView(doc, targetViewId);
                createdView.ViewTemplateId = ElementId.InvalidElementId;
                createdView.DetailLevel = ViewDetailLevel.Fine;
                createdView.DisplayStyle = DisplayStyle.ShadingWithEdges;
                createdView.AreModelCategoriesHidden = false;
                createdView.AreAnnotationCategoriesHidden = true;
                createdView.AreAnalyticalModelCategoriesHidden = true;
                createdView.AreImportCategoriesHidden = false;

                ApplyPhaseSettings(doc, createdView, log);
                ApplyViewParameters(createdView, settings.ViewParameters, log);
                ConfigureVisibilityGraphics(doc, createdView, log);

                tx.Commit();
            }
        }

        private static string GetRequestedViewName(string desiredName)
        {
            return string.IsNullOrWhiteSpace(desiredName) ? "KKY_CLEAN_3D" : desiredName.Trim();
        }

        private static bool ViewNameExists(Document doc, string viewName)
        {
            return new FilteredElementCollector(doc)
                .OfClass(typeof(View))
                .Cast<View>()
                .Any(x => string.Equals(x.Name, viewName, StringComparison.OrdinalIgnoreCase));
        }

        private static View3D GetRequiredTargetView(Document doc, ElementId targetViewId)
        {
            if (targetViewId == null || targetViewId == ElementId.InvalidElementId)
            {
                throw new InvalidOperationException("???3D 酉?ID媛 ?щ컮瑜댁? ?딆뒿?덈떎.");
            }

            View3D view = doc.GetElement(targetViewId) as View3D;
            if (view == null || !view.IsValidObject)
            {
                throw new InvalidOperationException("???3D 酉곌? ?좏슚?섏? ?딆뒿?덈떎. ?뺣━ 以???젣?섏뿀嫄곕굹 ?몃옖??뀡??濡ㅻ갚?섏뿀?듬땲??");
            }

            return view;
        }

        private static void ApplyPhaseSettings(Document doc, View view, Action<string> log)
        {
            Phase newConstruction = new FilteredElementCollector(doc)
                .OfClass(typeof(Phase))
                .Cast<Phase>()
                .FirstOrDefault(x => string.Equals(x.Name, "New Construction", StringComparison.OrdinalIgnoreCase)
                                  || string.Equals(x.Name, "?좎텞", StringComparison.OrdinalIgnoreCase));

            PhaseFilter showAll = new FilteredElementCollector(doc)
                .OfClass(typeof(PhaseFilter))
                .Cast<PhaseFilter>()
                .FirstOrDefault(x => string.Equals(x.Name, "Show All", StringComparison.OrdinalIgnoreCase)
                                  || string.Equals(x.Name, "紐⑤몢 ?쒖떆", StringComparison.OrdinalIgnoreCase));

            Parameter phaseParam = view.get_Parameter(BuiltInParameter.VIEW_PHASE);
            if (phaseParam != null && !phaseParam.IsReadOnly && newConstruction != null)
            {
                phaseParam.Set(newConstruction.Id);
                log?.Invoke("酉?Phase ?ㅼ젙: " + newConstruction.Name);
            }

            Parameter phaseFilterParam = view.get_Parameter(BuiltInParameter.VIEW_PHASE_FILTER);
            if (phaseFilterParam != null && !phaseFilterParam.IsReadOnly && showAll != null)
            {
                phaseFilterParam.Set(showAll.Id);
                log?.Invoke("酉?Phase Filter ?ㅼ젙: " + showAll.Name);
            }
        }

        private static void ApplyViewParameters(View view, IList<ViewParameterAssignment> assignments, Action<string> log)
        {
            if (assignments == null) return;

            foreach (ViewParameterAssignment item in assignments.Where(x => x != null && !string.IsNullOrWhiteSpace(x.ParameterName)))
            {
                Parameter parameter = view.LookupParameter(item.ParameterName);
                if (parameter == null || parameter.IsReadOnly)
                {
                    log?.Invoke("酉??뚮씪誘명꽣 誘몄쟻?? " + item.ParameterName);
                    continue;
                }

                try
                {
                    switch (parameter.StorageType)
                    {
                        case StorageType.String:
                            parameter.Set(item.ParameterValue ?? string.Empty);
                            break;
                        case StorageType.Integer:
                            parameter.Set(ParseInt(item.ParameterValue));
                            break;
                        case StorageType.Double:
                            parameter.Set(ParseDouble(item.ParameterValue));
                            break;
                        case StorageType.ElementId:
                            parameter.Set(new ElementId(ParseInt(item.ParameterValue)));
                            break;
                    }

                    log?.Invoke("酉??뚮씪誘명꽣 ?곸슜: " + item.ParameterName + " = " + item.ParameterValue);
                }
                catch (Exception ex)
                {
                    log?.Invoke("酉??뚮씪誘명꽣 ?곸슜 ?ㅽ뙣: " + item.ParameterName + " / " + ex.Message);
                }
            }
        }

        private static int ParseInt(string value)
        {
            int result;
            int.TryParse(value, out result);
            return result;
        }

        private static double ParseDouble(string value)
        {
            double result;
            double.TryParse(value, System.Globalization.NumberStyles.Float | System.Globalization.NumberStyles.AllowThousands,
                System.Globalization.CultureInfo.InvariantCulture, out result);
            return result;
        }

        private static void ConfigureVisibilityGraphics(Document doc, View view, Action<string> log)
        {
            int shownCount = 0;
            int hiddenCount = 0;
            int hiddenSubCount = 0;

            foreach (Category category in doc.Settings.Categories)
            {
                if (category == null) continue;
                if (!ShouldTreatAsModelVisibilityCategory(category)) continue;

                bool categoryVisible = !ShouldHideTopLevelCategory(category);
                if (TrySetCategoryVisibility(view, category, categoryVisible))
                {
                    if (categoryVisible) shownCount++;
                    else hiddenCount++;
                }

                foreach (Category subCategory in EnumerateAllSubCategories(category))
                {
                    if (subCategory == null) continue;

                    bool subVisible = categoryVisible;
                    if (ShouldHideNamedSubCategory(category, subCategory))
                    {
                        subVisible = false;
                    }

                    if (TrySetCategoryVisibility(view, subCategory, subVisible) && !subVisible)
                    {
                        hiddenSubCount++;
                    }
                }
            }

            log?.Invoke($"VG ?ㅼ젙 ?곸슜 ?꾨즺 / ?쒖떆 {shownCount}, ?④? {hiddenCount}, ?④? ?쒕툕移댄뀒怨좊━ {hiddenSubCount}");
        }

        private static bool ShouldTreatAsModelVisibilityCategory(Category category)
        {
            if (category.CategoryType == CategoryType.Model) return true;
            return IsLineCategory(category);
        }

        private static bool ShouldHideTopLevelCategory(Category category)
        {
            int id = category.Id.IntegerValue;
            return id == (int)BuiltInCategory.OST_Mass
                   || id == (int)BuiltInCategory.OST_Parts
                   || id == (int)BuiltInCategory.OST_Site
                   || id == (int)BuiltInCategory.OST_Lines
                   || EqualsNormalizedCategoryName(category, "Mass")
                   || EqualsNormalizedCategoryName(category, "Parts")
                   || EqualsNormalizedCategoryName(category, "Site")
                   || EqualsNormalizedCategoryName(category, "Lines")
                   || EqualsNormalizedCategoryName(category, "留ㅼ뒪")
                   || EqualsNormalizedCategoryName(category, "?뚯툩")
                   || EqualsNormalizedCategoryName(category, "?吏")
                   || EqualsNormalizedCategoryName(category, "선");
        }

        private static bool IsLineCategory(Category category)
        {
            int id = category.Id.IntegerValue;
            return id == (int)BuiltInCategory.OST_Lines
                   || EqualsNormalizedCategoryName(category, "Lines")
                   || EqualsNormalizedCategoryName(category, "선");
        }

        private static bool IsTargetFittingCategory(Category category)
        {
            int id = category.Id.IntegerValue;
            return id == (int)BuiltInCategory.OST_CableTrayFitting
                   || id == (int)BuiltInCategory.OST_ConduitFitting
                   || id == (int)BuiltInCategory.OST_DuctFitting
                   || id == (int)BuiltInCategory.OST_PipeFitting
                   || ContainsNormalizedCategoryName(category, "cablefitting")
                   || ContainsNormalizedCategoryName(category, "conduitfitting")
                   || ContainsNormalizedCategoryName(category, "ductfitting")
                   || ContainsNormalizedCategoryName(category, "pipefitting");
        }

        private static bool ShouldHideNamedSubCategory(Category parentCategory, Category subCategory)
        {
            if (!IsTargetFittingCategory(parentCategory)) return false;

            string name = NormalizeCategoryName(subCategory?.Name);
            if (string.IsNullOrWhiteSpace(name)) return false;

            return name.Contains("end") || name.Contains("cut");
        }

        private static IEnumerable<Category> EnumerateAllSubCategories(Category category)
        {
            var visited = new HashSet<int>();
            foreach (Category child in EnumerateAllSubCategoriesRecursive(category, visited))
            {
                yield return child;
            }
        }

        private static IEnumerable<Category> EnumerateAllSubCategoriesRecursive(Category category, HashSet<int> visited)
        {
            if (category == null || category.SubCategories == null) yield break;

            foreach (Category child in category.SubCategories)
            {
                if (child == null) continue;
                if (!visited.Add(child.Id.IntegerValue)) continue;

                yield return child;

                foreach (Category nested in EnumerateAllSubCategoriesRecursive(child, visited))
                {
                    yield return nested;
                }
            }
        }

        private static bool EqualsNormalizedCategoryName(Category category, string expected)
        {
            return string.Equals(NormalizeCategoryName(category?.Name), NormalizeCategoryName(expected), StringComparison.OrdinalIgnoreCase);
        }

        private static bool ContainsNormalizedCategoryName(Category category, string token)
        {
            string name = NormalizeCategoryName(category?.Name);
            string work = NormalizeCategoryName(token);
            return !string.IsNullOrWhiteSpace(name) && !string.IsNullOrWhiteSpace(work) && name.Contains(work);
        }

        private static string NormalizeCategoryName(string value)
        {
            if (string.IsNullOrWhiteSpace(value)) return string.Empty;
            return new string(value.Where(c => !char.IsWhiteSpace(c) && c != '-' && c != '_').ToArray()).ToLowerInvariant();
        }

        private static bool TrySetCategoryVisibility(View view, Category category, bool visible)
        {
            if (category == null) return false;
            if (!view.CanCategoryBeHidden(category.Id)) return false;
            view.SetCategoryHidden(category.Id, !visible);
            return true;
        }

        private static ElementId CreateOrApplyFilter(Document doc, ElementId targetViewId, BatchCleanSettings settings, Action<string> log)
        {
            ElementId filterId;
            using (var tx = new Transaction(doc, "Create/Apply View Filter"))
            {
                tx.Start();
                AttachFailureSwallower(tx);

                View3D view = GetRequiredTargetView(doc, targetViewId);
                filterId = RevitViewFilterProfileService.EnsureFilter(doc, settings.FilterProfile, log);
                RevitViewFilterProfileService.ApplyFilterToView(view, filterId, settings.ApplyFilterInitially);

                if (!settings.ApplyFilterInitially && settings.AutoEnableFilterIfEmpty)
                {
                    bool shouldEnable = RevitViewFilterProfileService.WouldViewBeEmptyWhenFilterHidden(doc, view, filterId);
                    if (shouldEnable)
                    {
                        RevitViewFilterProfileService.ApplyFilterToView(view, filterId, true);
                        log?.Invoke("?꾪꽣瑜?誘몄쟻???곹깭濡??먮㈃ 酉곌? 鍮꾩뼱 ?먮룞 ?쒖꽦?뷀뻽?듬땲??");
                    }
                    else
                    {
                        log?.Invoke("?꾪꽣瑜?誘몄쟻???곹깭濡??좎??덉뒿?덈떎. ?꾪꽣 ?놁씠??酉곗뿉 媛앹껜媛 ?쒖떆?⑸땲??");
                    }
                }

                tx.Commit();
            }

            return filterId;
        }

        private static void DeleteAllOtherViewsAndTemplates(Document doc, ElementId targetViewId, Action<string> log)
        {
            List<ElementId> deleteIds = new FilteredElementCollector(doc)
                .OfClass(typeof(View))
                .Cast<View>()
                .Where(x => x != null)
                .Where(x => targetViewId == null || x.Id.IntegerValue != targetViewId.IntegerValue)
                .Select(x => x.Id)
                .Distinct(new ElementIdComparer())
                .ToList();

            using (var tx = new Transaction(doc, "Delete Other Views and Templates"))
            {
                tx.Start();
                AttachFailureSwallower(tx);

                int deletedCount = 0;
                int failedCount = 0;
                foreach (ElementId viewId in deleteIds)
                {
                    if (viewId == null || viewId == ElementId.InvalidElementId)
                    {
                        continue;
                    }

                    try
                    {
                        if (targetViewId != null && viewId.IntegerValue == targetViewId.IntegerValue)
                        {
                            continue;
                        }

                        if (doc.GetElement(viewId) == null)
                        {
                            continue;
                        }

                        ICollection<ElementId> removed = doc.Delete(viewId);
                        if (removed != null && targetViewId != null && removed.Any(x => x != null && x.IntegerValue == targetViewId.IntegerValue))
                        {
                            throw new InvalidOperationException("???3D 酉곌? ?곗뇙 ??젣 ??곸쑝濡?媛먯??섏뿀?듬땲??");
                        }

                        deletedCount++;
                    }
                    catch
                    {
                        failedCount++;
                    }
                }

                tx.Commit();
                log?.Invoke($"??젣??酉??쒗뵆由? ???{deleteIds.Count}, ?깃났 {deletedCount}, ?ㅽ뙣 {failedCount}");
            }
        }

        private static void SetStartingView(Document doc, ElementId targetViewId, Action<string> log)
        {
            using (var tx = new Transaction(doc, "Set Starting View"))
            {
                tx.Start();
                AttachFailureSwallower(tx);
                try
                {
                    GetRequiredTargetView(doc, targetViewId);
                    StartingViewSettings startingViewSettings = StartingViewSettings.GetStartingViewSettings(doc);
                    startingViewSettings.ViewId = targetViewId;
                    tx.Commit();
                    log?.Invoke("Starting View ?ㅼ젙 ?꾨즺");
                }
                catch (Exception ex)
                {
                    tx.RollBack();
                    log?.Invoke("Starting View ?ㅼ젙 ?ㅽ뙣: " + ex.Message);
                }
            }
        }

        private static void DeleteUnusedViewFilters(Document doc, ElementId keepFilterId, Action<string> log)
        {
            using (var tx = new Transaction(doc, "Delete Unused View Filters"))
            {
                tx.Start();
                AttachFailureSwallower(tx);

                var usedFilterIds = new HashSet<int>();
                List<ElementId> viewIds = new FilteredElementCollector(doc)
                    .OfClass(typeof(View))
                    .ToElementIds()
                    .ToList();

                foreach (ElementId viewId in viewIds)
                {
                    try
                    {
                        View view = doc.GetElement(viewId) as View;
                        if (view == null || !view.IsValidObject)
                        {
                            continue;
                        }

                        foreach (ElementId filterId in view.GetFilters())
                        {
                            if (filterId != null)
                            {
                                usedFilterIds.Add(filterId.IntegerValue);
                            }
                        }
                    }
                    catch
                    {
                        // ignore
                    }
                }

                var deleteIds = new FilteredElementCollector(doc)
                    .OfClass(typeof(FilterElement))
                    .Cast<FilterElement>()
                    .Where(x => x != null)
                    .Where(x => keepFilterId == null || x.Id.IntegerValue != keepFilterId.IntegerValue)
                    .Where(x => !usedFilterIds.Contains(x.Id.IntegerValue))
                    .Select(x => x.Id)
                    .ToList();

                if (deleteIds.Count > 0)
                {
                    doc.Delete(deleteIds);
                }

                tx.Commit();
                log?.Invoke("誘몄궗??酉??꾪꽣 ??젣: " + deleteIds.Count);
            }
        }

        private static void TouchGroupsBrowser(Document doc, Action<string> log)
        {
            List<ElementId> candidateIds = new FilteredElementCollector(doc)
                .WhereElementIsNotElementType()
                .Where(x => x != null)
                .Where(x => !x.ViewSpecific)
                .Where(x => x.Category != null)
                .Where(x => x.Category.CategoryType == CategoryType.Model)
                .Where(x => !(x is Group))
                .Where(x => !(x is AssemblyInstance))
                .Select(x => x.Id)
                .Distinct(new ElementIdComparer())
                .Take(20)
                .ToList();

            if (candidateIds.Count == 0)
            {
                log?.Invoke("Groups 媛깆떊 ?좊룄 ?ㅽ궢: 洹몃９ ?앹꽦 ?꾨낫 ?붿냼 ?놁쓬");
                return;
            }

            using (var tx = new Transaction(doc, "Touch Groups Browser"))
            {
                tx.Start();
                AttachFailureSwallower(tx);

                foreach (ElementId candidateId in candidateIds)
                {
                    if (candidateId == null || candidateId == ElementId.InvalidElementId)
                    {
                        continue;
                    }

                    try
                    {
                        Element candidate = doc.GetElement(candidateId);
                        if (candidate == null || !candidate.IsValidObject)
                        {
                            continue;
                        }

                        Group group = doc.Create.NewGroup(new List<ElementId> { candidateId });
                        if (group == null || !group.IsValidObject)
                        {
                            continue;
                        }

                        ElementId groupTypeId = group.GroupType != null ? group.GroupType.Id : ElementId.InvalidElementId;

                        try
                        {
                            group.UngroupMembers();
                        }
                        catch
                        {
                            try
                            {
                                if (group.IsValidObject)
                                {
                                    doc.Delete(group.Id);
                                }
                            }
                            catch
                            {
                                // ignore
                            }
                        }

                        try
                        {
                            if (groupTypeId != null && groupTypeId != ElementId.InvalidElementId && doc.GetElement(groupTypeId) != null)
                            {
                                doc.Delete(groupTypeId);
                            }
                        }
                        catch
                        {
                            // ignore
                        }

                        tx.Commit();
                        log?.Invoke("Groups 媛깆떊 ?좊룄 ?꾨즺");
                        return;
                    }
                    catch
                    {
                        // try next candidate
                    }
                }

                tx.RollBack();
                log?.Invoke("Groups 媛깆떊 ?좊룄 ?ㅽ뙣: 洹몃９ ?앹꽦 媛?ν븳 紐⑤뜽 ?붿냼瑜?李얠? 紐삵뻽?듬땲??");
            }
        }

        private static void CollectDesignOptionAuditRows(Document doc, string sourcePath, List<DesignOptionAuditRow> rows, Action<string> log)
        {
            if (rows == null) return;

            List<DesignOption> options = new FilteredElementCollector(doc)
                .OfClass(typeof(DesignOption))
                .Cast<DesignOption>()
                .Where(x => x != null)
                .OrderBy(x => x.Name)
                .ToList();

            if (options.Count == 0)
            {
                rows.Add(new DesignOptionAuditRow
                {
                    SourcePath = sourcePath,
                    FileName = Path.GetFileName(sourcePath),
                    HasDesignOptions = "No",
                    OptionId = string.Empty,
                    OptionName = string.Empty,
                    IsPrimary = string.Empty,
                    MemberElementCount = string.Empty
                });

                log?.Invoke("Design Option ?놁쓬");
                return;
            }

            int totalOptionMembers = 0;
            foreach (DesignOption option in options)
            {
                int memberCount = 0;
                try
                {
                    memberCount = new FilteredElementCollector(doc)
                        .WhereElementIsNotElementType()
                        .WherePasses(new ElementDesignOptionFilter(option.Id))
                        .GetElementCount();
                }
                catch
                {
                    memberCount = 0;
                }

                totalOptionMembers += memberCount;

                rows.Add(new DesignOptionAuditRow
                {
                    SourcePath = sourcePath,
                    FileName = Path.GetFileName(sourcePath),
                    HasDesignOptions = "Yes",
                    OptionId = option.Id.IntegerValue.ToString(),
                    OptionName = option.Name ?? string.Empty,
                    IsPrimary = option.IsPrimary ? "Yes" : "No",
                    MemberElementCount = memberCount.ToString()
                });
            }

            log?.Invoke($"Design Option 諛쒓껄: ?듭뀡 {options.Count}媛?/ ?듭뀡 ?뚯냽 ?붿냼 {totalOptionMembers}媛?/ CSV 蹂닿퀬?쒕줈 ????덉젙");
        }

        private static string WriteDesignOptionAuditCsv(string outputFolder, IList<DesignOptionAuditRow> rows)
        {
            if (rows == null || rows.Count == 0) return null;
            if (string.IsNullOrWhiteSpace(outputFolder)) return null;

            Directory.CreateDirectory(outputFolder);

            string path = Path.Combine(outputFolder, "DesignOptionReport_" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + ".csv");
            var sb = new StringBuilder();
            sb.AppendLine("FileName,SourcePath,HasDesignOptions,OptionId,OptionName,IsPrimary,MemberElementCount");

            foreach (DesignOptionAuditRow row in rows)
            {
                sb.AppendLine(string.Join(",",
                    EscapeCsv(row.FileName),
                    EscapeCsv(row.SourcePath),
                    EscapeCsv(row.HasDesignOptions),
                    EscapeCsv(row.OptionId),
                    EscapeCsv(row.OptionName),
                    EscapeCsv(row.IsPrimary),
                    EscapeCsv(row.MemberElementCount)));
            }

            File.WriteAllText(path, sb.ToString(), new UTF8Encoding(true));
            return path;
        }

        private static ModelObjectCountComparison BuildCleanCountComparison(PreparedDocumentEntry prepared, bool success, string note)
        {
            if (prepared == null)
            {
                return BuildCleanCountComparison(string.Empty, null, success, note);
            }

            return new ModelObjectCountComparison
            {
                FileName = Path.GetFileName(prepared.OutputPath ?? prepared.SourcePath ?? string.Empty),
                SourcePath = prepared.SourcePath,
                OutputPath = prepared.OutputPath,
                BeforeCount = prepared.BeforeObjectCount,
                AfterCount = prepared.AfterObjectCount,
                Status = success ? "O" : "X",
                Note = string.IsNullOrWhiteSpace(note) ? (prepared.RemovedSummary ?? string.Empty) : note
            };
        }

        private static ModelObjectCountComparison BuildCleanCountComparison(string sourcePath, PreparedDocumentEntry prepared, bool success, string note)
        {
            return new ModelObjectCountComparison
            {
                FileName = Path.GetFileName((prepared != null ? prepared.OutputPath : null) ?? sourcePath ?? string.Empty),
                SourcePath = sourcePath,
                OutputPath = prepared?.OutputPath,
                BeforeCount = prepared != null ? (int?)prepared.BeforeObjectCount : null,
                AfterCount = prepared != null ? (int?)prepared.AfterObjectCount : null,
                Status = success ? "O" : "X",
                Note = note ?? string.Empty
            };
        }

        private static string EscapeCsv(string value)
        {
            string textValue = value ?? string.Empty;
            bool mustQuote = textValue.Contains(",") || textValue.Contains("\"") || textValue.Contains("\r") || textValue.Contains("\n");
            if (textValue.Contains("\""))
            {
                textValue = textValue.Replace("\"", "\"\"");
            }

            return mustQuote ? "\"" + textValue + "\"" : textValue;
        }
        private static void ApplyElementParameterUpdate(Document doc, ElementParameterUpdateSettings settings, Action<string> log)
        {
            if (doc == null) throw new ArgumentNullException(nameof(doc));
            if (settings == null || !settings.IsConfigured()) return;

            List<ElementParameterCondition> activeConditions = settings.Conditions
                .Where(x => x != null && x.IsConfigured())
                .ToList();
            List<ElementParameterAssignment> activeAssignments = settings.Assignments
                .Where(x => x != null && x.IsConfigured())
                .ToList();

            if (activeConditions.Count == 0 || activeAssignments.Count == 0)
            {
                log?.Invoke("객체 파라미터 일괄 입력 건너뜀: 유효한 조건 또는 입력 항목이 없습니다.");
                return;
            }

            using (var tx = new Transaction(doc, "Apply Element Parameter Update"))
            {
                tx.Start();
                AttachFailureSwallower(tx);

                int scannedCount = 0;
                int matchedCount = 0;
                int updatedElementCount = 0;
                int updatedValueCount = 0;
                int failedCount = 0;
                var updatedTypeKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

                IList<Element> elements = new FilteredElementCollector(doc)
                    .WhereElementIsNotElementType()
                    .Cast<Element>()
                    .Where(x => x != null)
                    .Where(x => !x.ViewSpecific)
                    .Where(x => x.Category != null)
                    .Where(x => x.Category.CategoryType == CategoryType.Model)
                    .ToList();

                foreach (Element element in elements)
                {
                    scannedCount++;
                    try
                    {
                        if (!DoesElementMatchConditions(doc, element, activeConditions, settings.CombinationMode))
                        {
                            continue;
                        }

                        matchedCount++;
                        bool elementUpdated = false;

                        foreach (ElementParameterAssignment assignment in activeAssignments)
                        {
                            IList<ParameterTargetInfo> targets = FindParameterTargetsOnElementOrType(
                                doc,
                                element,
                                assignment.ParameterName,
                                true,
                                settings.ApplyToAllMatchingParameters);
                            if (targets == null || targets.Count == 0)
                            {
                                failedCount++;
                                continue;
                            }

                            foreach (ParameterTargetInfo target in targets)
                            {
                                if (target == null || target.Parameter == null || target.Owner == null)
                                {
                                    failedCount++;
                                    continue;
                                }

                                if (target.IsTypeParameter)
                                {
                                    string typeKey = target.Owner.Id.IntegerValue.ToString(CultureInfo.InvariantCulture)
                                        + "|" + (assignment.ParameterName ?? string.Empty)
                                        + "|" + (assignment.Value ?? string.Empty)
                                        + "|" + target.MatchIndex.ToString(CultureInfo.InvariantCulture);
                                    if (!updatedTypeKeys.Add(typeKey))
                                    {
                                        continue;
                                    }
                                }

                                if (TrySetParameterValue(doc, target.Owner, target.Parameter, assignment.Value, log))
                                {
                                    updatedValueCount++;
                                    elementUpdated = true;
                                }
                                else
                                {
                                    failedCount++;
                                }
                            }
                        }

                        if (elementUpdated)
                        {
                            updatedElementCount++;
                        }
                    }
                    catch
                    {
                        failedCount++;
                    }
                }

                tx.Commit();
                string duplicateModeText = settings.ApplyToAllMatchingParameters ? "중복 전체 입력" : "중복 하나만 입력";
                log?.Invoke($"객체 파라미터 일괄 입력: 스캔 {scannedCount}, 조건일치 {matchedCount}, 변경 요소 {updatedElementCount}, 값 입력 {updatedValueCount}, 실패 {failedCount}, 중복 처리 {duplicateModeText}");
            }
        }


        private sealed class ParameterTargetInfo
        {
            public Parameter Parameter { get; set; }
            public Element Owner { get; set; }
            public bool IsTypeParameter { get; set; }
            public int MatchIndex { get; set; }
        }

        private static IList<ParameterTargetInfo> FindParameterTargetsOnElementOrType(Document doc, Element element, string parameterName, bool requireWritable, bool applyAllMatchingParameters)
        {
            var results = new List<ParameterTargetInfo>();
            if (doc == null || element == null || string.IsNullOrWhiteSpace(parameterName))
            {
                return results;
            }

            List<Parameter> instanceParameters = FindParametersByName(element, parameterName, requireWritable);
            if (applyAllMatchingParameters)
            {
                for (int i = 0; i < instanceParameters.Count; i++)
                {
                    results.Add(new ParameterTargetInfo
                    {
                        Parameter = instanceParameters[i],
                        Owner = element,
                        IsTypeParameter = false,
                        MatchIndex = i
                    });
                }
            }
            else if (instanceParameters.Count > 0)
            {
                results.Add(new ParameterTargetInfo
                {
                    Parameter = instanceParameters[0],
                    Owner = element,
                    IsTypeParameter = false,
                    MatchIndex = 0
                });
                return results;
            }

            ElementId typeId = element.GetTypeId();
            if (typeId == null || typeId == ElementId.InvalidElementId)
            {
                return results;
            }

            Element typeElement = doc.GetElement(typeId);
            if (typeElement == null)
            {
                return results;
            }

            List<Parameter> typeParameters = FindParametersByName(typeElement, parameterName, requireWritable);
            if (applyAllMatchingParameters)
            {
                for (int i = 0; i < typeParameters.Count; i++)
                {
                    results.Add(new ParameterTargetInfo
                    {
                        Parameter = typeParameters[i],
                        Owner = typeElement,
                        IsTypeParameter = true,
                        MatchIndex = i
                    });
                }
            }
            else if (results.Count == 0 && typeParameters.Count > 0)
            {
                results.Add(new ParameterTargetInfo
                {
                    Parameter = typeParameters[0],
                    Owner = typeElement,
                    IsTypeParameter = true,
                    MatchIndex = 0
                });
            }

            return results;
        }

        private static Parameter FindParameterOnElementOrType(Document doc, Element element, string parameterName, bool requireWritable, out bool isTypeParameter, out Element owner)
        {
            ParameterTargetInfo target = FindParameterTargetsOnElementOrType(doc, element, parameterName, requireWritable, false).FirstOrDefault();
            isTypeParameter = target != null && target.IsTypeParameter;
            owner = target?.Owner;
            return target?.Parameter;
        }

        private static Parameter FindParameterByName(Element element, string parameterName, bool requireWritable)
        {
            return FindParametersByName(element, parameterName, requireWritable).FirstOrDefault();
        }

        private static List<Parameter> FindParametersByName(Element element, string parameterName, bool requireWritable)
        {
            var matches = new List<Parameter>();
            if (element == null || string.IsNullOrWhiteSpace(parameterName))
            {
                return matches;
            }

            Parameter direct = null;
            try
            {
                direct = element.LookupParameter(parameterName);
            }
            catch
            {
                direct = null;
            }

            if (IsUsableNamedParameter(direct, parameterName, requireWritable))
            {
                matches.Add(direct);
            }

            foreach (Parameter parameter in element.Parameters.Cast<Parameter>())
            {
                if (!IsUsableNamedParameter(parameter, parameterName, requireWritable))
                {
                    continue;
                }

                if (matches.Any(existing => ReferenceEquals(existing, parameter)))
                {
                    continue;
                }

                if (!requireWritable && parameter.HasValue)
                {
                    matches.Insert(0, parameter);
                    continue;
                }

                matches.Add(parameter);
            }

            return matches;
        }

        private static bool IsUsableNamedParameter(Parameter parameter, string parameterName, bool requireWritable)
        {
            if (parameter == null)
            {
                return false;
            }

            Definition definition = parameter.Definition;
            if (definition == null)
            {
                return false;
            }

            if (!string.Equals(definition.Name, parameterName, StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }

            if (requireWritable && parameter.IsReadOnly)
            {
                return false;
            }

            return true;
        }
        private static bool DoesElementMatchConditions(Document doc, Element element, IList<ElementParameterCondition> conditions, ParameterConditionCombination combinationMode)
        {
            if (conditions == null || conditions.Count == 0) return false;

            bool hasAny = false;
            bool anyMatched = false;
            foreach (ElementParameterCondition condition in conditions)
            {
                if (condition == null || !condition.IsConfigured())
                {
                    continue;
                }

                hasAny = true;
                Parameter filterParameter = FindParameterOnElementOrType(doc, element, condition.ParameterName, false, out bool isTypeParameter, out Element owner);
                bool matched = filterParameter != null && DoesParameterMatch(filterParameter, condition.Operator, condition.Value);

                if (combinationMode == ParameterConditionCombination.And && !matched)
                {
                    return false;
                }

                if (combinationMode == ParameterConditionCombination.Or && matched)
                {
                    anyMatched = true;
                }
            }

            if (!hasAny) return false;
            return combinationMode == ParameterConditionCombination.Or ? anyMatched : true;
        }

        private static bool DoesParameterMatch(Parameter parameter, FilterRuleOperator op, string expectedValue)
        {
            if (parameter == null) return false;

            string textValue = GetComparableParameterText(parameter);
            bool hasValue = !string.IsNullOrWhiteSpace(textValue);
            string ruleValue = expectedValue ?? string.Empty;

            if (parameter.StorageType == StorageType.Integer || parameter.StorageType == StorageType.Double)
            {
                double actualNumber;
                double expectedNumber;
                bool canActual = TryGetNumericParameterValue(parameter, out actualNumber);
                bool canExpected = TryParseNumber(ruleValue, out expectedNumber);

                switch (op)
                {
                    case FilterRuleOperator.HasValue:
                        return hasValue;
                    case FilterRuleOperator.HasNoValue:
                        return !hasValue;
                    case FilterRuleOperator.Greater:
                        return canActual && canExpected && actualNumber > expectedNumber;
                    case FilterRuleOperator.GreaterOrEqual:
                        return canActual && canExpected && actualNumber >= expectedNumber;
                    case FilterRuleOperator.Less:
                        return canActual && canExpected && actualNumber < expectedNumber;
                    case FilterRuleOperator.LessOrEqual:
                        return canActual && canExpected && actualNumber <= expectedNumber;
                    case FilterRuleOperator.Equals:
                        return canActual && canExpected ? Math.Abs(actualNumber - expectedNumber) < 0.000001d : string.Equals(textValue, ruleValue, StringComparison.OrdinalIgnoreCase);
                    case FilterRuleOperator.NotEquals:
                        return canActual && canExpected ? Math.Abs(actualNumber - expectedNumber) >= 0.000001d : !string.Equals(textValue, ruleValue, StringComparison.OrdinalIgnoreCase);
                }
            }

            switch (op)
            {
                case FilterRuleOperator.Equals:
                    return string.Equals(textValue, ruleValue, StringComparison.OrdinalIgnoreCase);
                case FilterRuleOperator.NotEquals:
                    return !string.Equals(textValue, ruleValue, StringComparison.OrdinalIgnoreCase);
                case FilterRuleOperator.Contains:
                    return textValue.IndexOf(ruleValue, StringComparison.OrdinalIgnoreCase) >= 0;
                case FilterRuleOperator.NotContains:
                    return textValue.IndexOf(ruleValue, StringComparison.OrdinalIgnoreCase) < 0;
                case FilterRuleOperator.BeginsWith:
                    return textValue.StartsWith(ruleValue, StringComparison.OrdinalIgnoreCase);
                case FilterRuleOperator.NotBeginsWith:
                    return !textValue.StartsWith(ruleValue, StringComparison.OrdinalIgnoreCase);
                case FilterRuleOperator.EndsWith:
                    return textValue.EndsWith(ruleValue, StringComparison.OrdinalIgnoreCase);
                case FilterRuleOperator.NotEndsWith:
                    return !textValue.EndsWith(ruleValue, StringComparison.OrdinalIgnoreCase);
                case FilterRuleOperator.HasValue:
                    return hasValue;
                case FilterRuleOperator.HasNoValue:
                    return !hasValue;
                case FilterRuleOperator.Greater:
                case FilterRuleOperator.GreaterOrEqual:
                case FilterRuleOperator.Less:
                case FilterRuleOperator.LessOrEqual:
                    double actual;
                    double expected;
                    return TryParseNumber(textValue, out actual) && TryParseNumber(ruleValue, out expected) && CompareNumbers(actual, expected, op);
                default:
                    return false;
            }
        }

        private static bool CompareNumbers(double actual, double expected, FilterRuleOperator op)
        {
            switch (op)
            {
                case FilterRuleOperator.Greater:
                    return actual > expected;
                case FilterRuleOperator.GreaterOrEqual:
                    return actual >= expected;
                case FilterRuleOperator.Less:
                    return actual < expected;
                case FilterRuleOperator.LessOrEqual:
                    return actual <= expected;
                default:
                    return false;
            }
        }

        private static bool TrySetParameterValue(Document doc, Element owner, Parameter parameter, string value, Action<string> log)
        {
            if (parameter == null || parameter.IsReadOnly) return false;

            try
            {
                switch (parameter.StorageType)
                {
                    case StorageType.String:
                        parameter.Set(value ?? string.Empty);
                        return true;
                    case StorageType.Integer:
                        if (int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int intValue) || int.TryParse(value, out intValue))
                        {
                            parameter.Set(intValue);
                            return true;
                        }
                        if (bool.TryParse(value, out bool boolValue))
                        {
                            parameter.Set(boolValue ? 1 : 0);
                            return true;
                        }
                        return false;
                    case StorageType.Double:
                        if (double.TryParse(value, NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.InvariantCulture, out double dblValue) || double.TryParse(value, out dblValue))
                        {
                            parameter.Set(dblValue);
                            return true;
                        }
                        return false;
                    case StorageType.ElementId:
                        if (int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int idValue) || int.TryParse(value, out idValue))
                        {
                            parameter.Set(new ElementId(idValue));
                            return true;
                        }
                        return false;
                    default:
                        return false;
                }
            }
            catch (Exception ex)
            {
                log?.Invoke("?뚮씪誘명꽣 媛??낅젰 ?ㅽ뙣: " + owner?.Id?.IntegerValue + " / " + parameter.Definition?.Name + " / " + ex.Message);
                return false;
            }
        }

        private static string GetComparableParameterText(Parameter parameter)
        {
            if (parameter == null) return string.Empty;

            try
            {
                switch (parameter.StorageType)
                {
                    case StorageType.String:
                        return parameter.AsString() ?? parameter.AsValueString() ?? string.Empty;
                    case StorageType.Integer:
                        return parameter.AsValueString() ?? parameter.AsInteger().ToString(CultureInfo.InvariantCulture);
                    case StorageType.Double:
                        return parameter.AsValueString() ?? parameter.AsDouble().ToString(CultureInfo.InvariantCulture);
                    case StorageType.ElementId:
                        ElementId id = parameter.AsElementId();
                        if (id == null || id == ElementId.InvalidElementId) return string.Empty;
                        return id.IntegerValue.ToString(CultureInfo.InvariantCulture);
                    default:
                        return parameter.AsValueString() ?? string.Empty;
                }
            }
            catch
            {
                return string.Empty;
            }
        }

        private static bool TryGetNumericParameterValue(Parameter parameter, out double value)
        {
            value = 0d;
            if (parameter == null) return false;

            try
            {
                switch (parameter.StorageType)
                {
                    case StorageType.Integer:
                        value = parameter.AsInteger();
                        return true;
                    case StorageType.Double:
                        value = parameter.AsDouble();
                        return true;
                    default:
                        return TryParseNumber(GetComparableParameterText(parameter), out value);
                }
            }
            catch
            {
                return false;
            }
        }

        private static bool TryParseNumber(string text, out double value)
        {
            if (double.TryParse(text, NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.InvariantCulture, out value))
            {
                return true;
            }

            return double.TryParse(text, out value);
        }

        private static string BuildOutputPath(string sourcePath, string outputFolder)
        {
            string fileName = Path.GetFileName(sourcePath);
            if (fileName.IndexOf("_Detached", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                fileName = fileName.Replace("_Detached", string.Empty).Replace("_detached", string.Empty);
            }

            return Path.Combine(outputFolder, fileName);
        }

        private static void SaveCleanDocument(Document doc, ElementId targetViewId, string outputPath, Action<string> log)
        {
            using (var tx = new Transaction(doc, "Finalize Save Settings"))
            {
                tx.Start();
                AttachFailureSwallower(tx);
                try
                {
                    View3D targetView = GetRequiredTargetView(doc, targetViewId);
                    targetView.ViewTemplateId = ElementId.InvalidElementId;
                    tx.Commit();
                }
                catch
                {
                    tx.RollBack();
                }
            }

            View3D saveView = GetRequiredTargetView(doc, targetViewId);
            string previewViewName = saveView.Name;
            SaveAsOptions saveAsOptions = new SaveAsOptions
            {
                OverwriteExistingFile = true,
                Compact = true,
                MaximumBackups = 1
            };

            doc.SaveAs(outputPath, saveAsOptions);
            log?.Invoke("????꾨즺: " + outputPath + " / TargetView=" + previewViewName);
        }

        private static void AttachFailureSwallower(Transaction tx)
        {
            FailureHandlingOptions options = tx.GetFailureHandlingOptions();
            options.SetFailuresPreprocessor(new TransactionFailureSwallower());
            tx.SetFailureHandlingOptions(options);
        }

        private sealed class DesignOptionAuditRow
        {
            public string SourcePath { get; set; }
            public string FileName { get; set; }
            public string HasDesignOptions { get; set; }
            public string OptionId { get; set; }
            public string OptionName { get; set; }
            public string IsPrimary { get; set; }
            public string MemberElementCount { get; set; }
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
