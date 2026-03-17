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
                    log?.Invoke("정리 시작: " + filePath);

                    PreparedDocumentEntry prepared = PrepareSingle(uiapp, filePath, settings, designOptionAuditRows, log);
                    session.PreparedDocuments.Add(prepared);
                    result.OutputPath = prepared.OutputPath;
                    result.Success = true;
                    result.Message = "정리 완료 / 저장 대기";

                    log?.Invoke("정리 완료(저장 대기): " + prepared.OutputPath);
                }
                catch (Exception ex)
                {
                    result.Success = false;
                    result.Message = ex.Message;
                    log?.Invoke("실패: " + ex.Message);
                }
            }

            try
            {
                string auditPath = WriteDesignOptionAuditCsv(settings.OutputFolder, designOptionAuditRows);
                session.DesignOptionAuditCsvPath = auditPath;
                if (!string.IsNullOrWhiteSpace(auditPath))
                {
                    log?.Invoke("Design Option CSV 저장: " + auditPath);
                }
            }
            catch (Exception ex)
            {
                log?.Invoke("Design Option CSV 저장 실패: " + ex.Message);
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
                    log?.Invoke("정리 시작: " + filePath);

                    prepared = PrepareSingle(uiapp, filePath, settings, designOptionAuditRows, log);
                    result.OutputPath = prepared.OutputPath;

                    log?.Invoke("[STEP] Groups 브라우저 갱신 유도 시작");
                    try
                    {
                        TouchGroupsBrowser(prepared.Document, log);
                    }
                    catch (Exception ex)
                    {
                        log?.Invoke("Groups 갱신 유도 실패(무시): " + ex.Message);
                    }

                    log?.Invoke("[STEP] 저장 시작");
                    SaveCleanDocument(prepared.Document, prepared.TargetViewId, prepared.OutputPath, log);
                    session.CleanedOutputPaths.Add(prepared.OutputPath);

                    result.Success = true;
                    result.Message = "정리 및 저장 완료";
                    log?.Invoke("정리 및 저장 완료: " + prepared.OutputPath);
                }
                catch (Exception ex)
                {
                    result.Success = false;
                    result.Message = ex.Message;
                    log?.Invoke("실패: " + ex.Message);
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
                    log?.Invoke("Design Option CSV 저장: " + auditPath);
                }
            }
            catch (Exception ex)
            {
                log?.Invoke("Design Option CSV 저장 실패: " + ex.Message);
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
                        throw new InvalidOperationException("저장 대상 문서가 유효하지 않습니다.");
                    }

                    log?.Invoke("----------------------------------------");
                    log?.Invoke("저장 시작: " + prepared.SourcePath);
                    log?.Invoke("[STEP] Groups 브라우저 갱신 유도 시작");
                    try
                    {
                        TouchGroupsBrowser(doc, log);
                    }
                    catch (Exception ex)
                    {
                        log?.Invoke("Groups 갱신 유도 실패(무시): " + ex.Message);
                    }

                    log?.Invoke("[STEP] 저장 시작");
                    SaveCleanDocument(doc, prepared.TargetViewId, prepared.OutputPath, log);
                    result.Success = true;
                    result.Message = "저장 완료";
                }
                catch (Exception ex)
                {
                    result.Success = false;
                    result.Message = ex.Message;
                    log?.Invoke("저장 실패: " + ex.Message);
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
                        log?.Invoke("저장 없이 닫음: " + prepared.SourcePath);
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
                if (doc == null) throw new InvalidOperationException("문서를 열지 못했습니다.");

                CollectDesignOptionAuditRows(doc, filePath, designOptionAuditRows, log);

                log?.Invoke("[STEP] Manage Links 정리 시작");
                DeleteManageLinks(doc, log);
                log?.Invoke("[STEP] Group/Assembly 해제 시작");
                UngroupAndDisassemble(doc, log);
                log?.Invoke("[STEP] 기존 3D 뷰/동명 뷰 삭제 시작");
                DeleteExisting3DViewsAndConflictingNamedViews(doc, settings.Target3DViewName, log);

                log?.Invoke("[STEP] 대상 3D 뷰 생성 시작");
                ElementId targetViewId = CreateTarget3DView(doc, settings.Target3DViewName, log);
                log?.Invoke("[STEP] 대상 3D 뷰 설정 적용 시작");
                ConfigureTarget3DView(doc, targetViewId, settings, log);

                ElementId keptFilterId = ElementId.InvalidElementId;
                if (settings.UseFilter && settings.FilterProfile != null && settings.FilterProfile.IsConfigured())
                {
                    keptFilterId = CreateOrApplyFilter(doc, targetViewId, settings, log);
                }

                log?.Invoke("[STEP] 대상 뷰 외 나머지 뷰/템플릿 삭제 시작");
                DeleteAllOtherViewsAndTemplates(doc, targetViewId, log);
                log?.Invoke("[STEP] Starting View 설정 시작");
                SetStartingView(doc, targetViewId, log);
                log?.Invoke("[STEP] 미사용 뷰 필터 삭제 시작");
                DeleteUnusedViewFilters(doc, keptFilterId, log);

                if (settings.ElementParameterUpdate != null && settings.ElementParameterUpdate.IsConfigured())
                {
                    log?.Invoke("[STEP] 객체 파라미터 일괄 입력 시작");
                    ApplyElementParameterUpdate(doc, settings.ElementParameterUpdate, log);
                }
                else
                {
                    log?.Invoke("[STEP] 객체 파라미터 일괄 입력 건너뜀");
                }

                log?.Invoke("[STEP] Legacy Purge 제거됨 - 정리 단계에서는 Purge를 실행하지 않음");

                return new PreparedDocumentEntry
                {
                    SourcePath = filePath,
                    OutputPath = BuildOutputPath(filePath, settings.OutputFolder),
                    Document = doc,
                    TargetViewId = targetViewId,
                    KeptFilterId = keptFilterId
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

                log?.Invoke("Detach + 모든 워크셋 닫기 + workset discard 열기 시도");
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
                    log?.Invoke("Manage Links/Imports 삭제 결과 / 대상 " + ids.Count + ", 성공 " + deletedCount + ", 실패 " + failedDeleteCount);
                }

                tx.Commit();
                log?.Invoke("Manage Links/Imports 단계 완료");
                log?.Invoke("Manage Links/Imports 삭제 대상: " + ids.Count);
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
                        log?.Invoke(label + " 삭제 실패: Id " + id.IntegerValue + " / " + ex.Message);
                        detailedFailureLogCount++;
                    }
                }
            }

            failedCount += localFailedCount;
            log?.Invoke(label + " 삭제 시도: 대상 " + targets.Count + ", 성공 " + deletedCount + ", 실패 " + localFailedCount);
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
                log?.Invoke($"그룹 해제: {groupUngroupCount}, 그룹 삭제: {groupDeleteCount + leftoverGroupDeleteCount}, 그룹 타입 삭제: {groupTypeDeleteCount}, 어셈블리 해제: {assemblyDisassembleCount}, 어셈블리 삭제: {assemblyDeleteCount + leftoverAssemblyDeleteCount}, 어셈블리 타입 삭제: {assemblyTypeDeleteCount}");
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
                throw new InvalidOperationException("사용자 지정 3D 뷰 이름이 기존 뷰와 충돌하여 정확한 이름으로 만들 수 없습니다: " + exactName);
            }

            log?.Invoke($"기존 3D 뷰/동명 뷰 선삭제: 대상 {candidateIds.Count}, 성공 {deletedCount}, 실패 {failedCount}");
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
                    throw new InvalidOperationException("대상 3D 뷰 이름이 아직 문서에 남아 있습니다: " + exactName);
                }

                ViewFamilyType threeDType = new FilteredElementCollector(doc)
                    .OfClass(typeof(ViewFamilyType))
                    .Cast<ViewFamilyType>()
                    .FirstOrDefault(x => x.ViewFamily == ViewFamily.ThreeDimensional);

                if (threeDType == null)
                {
                    throw new InvalidOperationException("3D ViewFamilyType을 찾지 못했습니다.");
                }

                View3D createdView = View3D.CreateIsometric(doc, threeDType.Id);
                createdView.Name = exactName;
                createdView.ViewTemplateId = ElementId.InvalidElementId;
                createdViewId = createdView.Id;

                tx.Commit();
            }

            log?.Invoke("대상 3D 뷰 생성: " + exactName);
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
                throw new InvalidOperationException("대상 3D 뷰 ID가 올바르지 않습니다.");
            }

            View3D view = doc.GetElement(targetViewId) as View3D;
            if (view == null || !view.IsValidObject)
            {
                throw new InvalidOperationException("대상 3D 뷰가 유효하지 않습니다. 정리 중 삭제되었거나 트랜잭션이 롤백되었습니다.");
            }

            return view;
        }

        private static void ApplyPhaseSettings(Document doc, View view, Action<string> log)
        {
            Phase newConstruction = new FilteredElementCollector(doc)
                .OfClass(typeof(Phase))
                .Cast<Phase>()
                .FirstOrDefault(x => string.Equals(x.Name, "New Construction", StringComparison.OrdinalIgnoreCase)
                                  || string.Equals(x.Name, "신축", StringComparison.OrdinalIgnoreCase));

            PhaseFilter showAll = new FilteredElementCollector(doc)
                .OfClass(typeof(PhaseFilter))
                .Cast<PhaseFilter>()
                .FirstOrDefault(x => string.Equals(x.Name, "Show All", StringComparison.OrdinalIgnoreCase)
                                  || string.Equals(x.Name, "모두 표시", StringComparison.OrdinalIgnoreCase));

            Parameter phaseParam = view.get_Parameter(BuiltInParameter.VIEW_PHASE);
            if (phaseParam != null && !phaseParam.IsReadOnly && newConstruction != null)
            {
                phaseParam.Set(newConstruction.Id);
                log?.Invoke("뷰 Phase 설정: " + newConstruction.Name);
            }

            Parameter phaseFilterParam = view.get_Parameter(BuiltInParameter.VIEW_PHASE_FILTER);
            if (phaseFilterParam != null && !phaseFilterParam.IsReadOnly && showAll != null)
            {
                phaseFilterParam.Set(showAll.Id);
                log?.Invoke("뷰 Phase Filter 설정: " + showAll.Name);
            }
        }

        private static void ApplyViewParameters(View view, IList<ViewParameterAssignment> assignments, Action<string> log)
        {
            if (assignments == null) return;

            foreach (ViewParameterAssignment item in assignments.Where(x => x != null && x.Enabled && !string.IsNullOrWhiteSpace(x.ParameterName)))
            {
                Parameter parameter = view.LookupParameter(item.ParameterName);
                if (parameter == null || parameter.IsReadOnly)
                {
                    log?.Invoke("뷰 파라미터 미적용: " + item.ParameterName);
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

                    log?.Invoke("뷰 파라미터 적용: " + item.ParameterName + " = " + item.ParameterValue);
                }
                catch (Exception ex)
                {
                    log?.Invoke("뷰 파라미터 적용 실패: " + item.ParameterName + " / " + ex.Message);
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

            log?.Invoke($"VG 설정 적용 완료 / 표시 {shownCount}, 숨김 {hiddenCount}, 숨김 서브카테고리 {hiddenSubCount}");
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
                   || EqualsNormalizedCategoryName(category, "매스")
                   || EqualsNormalizedCategoryName(category, "파츠")
                   || EqualsNormalizedCategoryName(category, "대지")
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
                        log?.Invoke("필터를 미적용 상태로 두면 뷰가 비어 자동 활성화했습니다.");
                    }
                    else
                    {
                        log?.Invoke("필터를 미적용 상태로 유지했습니다. 필터 없이도 뷰에 객체가 표시됩니다.");
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
                            throw new InvalidOperationException("대상 3D 뷰가 연쇄 삭제 대상으로 감지되었습니다.");
                        }

                        deletedCount++;
                    }
                    catch
                    {
                        failedCount++;
                    }
                }

                tx.Commit();
                log?.Invoke($"삭제된 뷰/템플릿: 대상 {deleteIds.Count}, 성공 {deletedCount}, 실패 {failedCount}");
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
                    log?.Invoke("Starting View 설정 완료");
                }
                catch (Exception ex)
                {
                    tx.RollBack();
                    log?.Invoke("Starting View 설정 실패: " + ex.Message);
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
                log?.Invoke("미사용 뷰 필터 삭제: " + deleteIds.Count);
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
                log?.Invoke("Groups 갱신 유도 스킵: 그룹 생성 후보 요소 없음");
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
                        log?.Invoke("Groups 갱신 유도 완료");
                        return;
                    }
                    catch
                    {
                        // try next candidate
                    }
                }

                tx.RollBack();
                log?.Invoke("Groups 갱신 유도 실패: 그룹 생성 가능한 모델 요소를 찾지 못했습니다.");
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

                log?.Invoke("Design Option 없음");
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

            log?.Invoke($"Design Option 발견: 옵션 {options.Count}개 / 옵션 소속 요소 {totalOptionMembers}개 / CSV 보고서로 저장 예정");
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
                            Parameter targetParameter = FindParameterOnElementOrType(doc, element, assignment.ParameterName, true, out bool targetOnType, out Element targetOwner);
                            if (targetParameter == null || targetOwner == null)
                            {
                                failedCount++;
                                continue;
                            }

                            if (targetOnType)
                            {
                                string typeKey = targetOwner.Id.IntegerValue.ToString(CultureInfo.InvariantCulture)
                                    + "|" + (assignment.ParameterName ?? string.Empty)
                                    + "|" + (assignment.Value ?? string.Empty);
                                if (!updatedTypeKeys.Add(typeKey))
                                {
                                    continue;
                                }
                            }

                            if (TrySetParameterValue(doc, targetOwner, targetParameter, assignment.Value, log))
                            {
                                updatedValueCount++;
                                elementUpdated = true;
                            }
                            else
                            {
                                failedCount++;
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
                log?.Invoke($"객체 파라미터 일괄 입력: 스캔 {scannedCount}, 조건일치 {matchedCount}, 변경 요소 {updatedElementCount}, 값 입력 {updatedValueCount}, 실패 {failedCount}");
            }
        }


        private static Parameter FindParameterOnElementOrType(Document doc, Element element, string parameterName, bool requireWritable, out bool isTypeParameter, out Element owner)
        {
            isTypeParameter = false;
            owner = null;

            if (doc == null || element == null || string.IsNullOrWhiteSpace(parameterName))
            {
                return null;
            }

            Parameter instanceParameter = FindParameterByName(element, parameterName, requireWritable);
            if (instanceParameter != null)
            {
                owner = element;
                return instanceParameter;
            }

            ElementId typeId = element.GetTypeId();
            if (typeId == null || typeId == ElementId.InvalidElementId)
            {
                return null;
            }

            Element typeElement = doc.GetElement(typeId);
            if (typeElement == null)
            {
                return null;
            }

            Parameter typeParameter = FindParameterByName(typeElement, parameterName, requireWritable);
            if (typeParameter != null)
            {
                isTypeParameter = true;
                owner = typeElement;
                return typeParameter;
            }

            return null;
        }

        private static Parameter FindParameterByName(Element element, string parameterName, bool requireWritable)
        {
            if (element == null || string.IsNullOrWhiteSpace(parameterName))
            {
                return null;
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
                return direct;
            }

            Parameter bestMatch = null;
            foreach (Parameter parameter in element.Parameters.Cast<Parameter>())
            {
                if (!IsUsableNamedParameter(parameter, parameterName, requireWritable))
                {
                    continue;
                }

                if (!requireWritable && parameter.HasValue)
                {
                    return parameter;
                }

                if (bestMatch == null)
                {
                    bestMatch = parameter;
                }
            }

            return bestMatch;
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
                log?.Invoke("파라미터 값 입력 실패: " + owner?.Id?.IntegerValue + " / " + parameter.Definition?.Name + " / " + ex.Message);
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
            log?.Invoke("저장 완료: " + outputPath + " / TargetView=" + previewViewName);
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
