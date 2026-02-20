using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace KKY_Tool_Revit.Services
{
    public class FamilyLinkTargetParam
    {
        public string Name { get; set; } = "";
        public Guid Guid { get; set; }
        public string GroupName { get; set; } = "";
        public string DataTypeToken { get; set; } = "";
    }

    public class FamilyLinkAuditRow
    {
        public string FileName { get; set; } = "";
        public string HostFamilyName { get; set; } = "";
        public string HostFamilyCategory { get; set; } = "";
        public string NestedFamilyName { get; set; } = "";
        public string NestedTypeName { get; set; } = "";
        public string NestedCategory { get; set; } = "";
        public string NestedParamName { get; set; } = "";
        public string TargetParamName { get; set; } = "";
        public string ExpectedGuid { get; set; } = "";
        public string FoundScope { get; set; } = "";
        public string NestedParamGuid { get; set; } = "";
        public string NestedParamDataType { get; set; } = "";
        public string AssocHostParamName { get; set; } = "";
        public string HostParamGuid { get; set; } = "";
        public string HostParamIsShared { get; set; } = "";
        public string Issue { get; set; } = "";
        public string Notes { get; set; } = "";
    }

    internal enum FamilyLinkAuditIssue
    {
        OK,
        MissingAssociation,
        GuidMismatch,
        HostParamNotShared,
        ParamNotFound,
        Error
    }

    internal enum FoundScope
    {
        InstanceParam,
        TypeParam
    }

    internal class FoundParam
    {
        public Parameter P { get; set; }
        public FoundScope Scope { get; set; }
    }

    public static class FamilyLinkAuditService
    {
        public static List<FamilyLinkAuditRow> Run(
            UIApplication app,
            IList<string> rvtPaths,
            IList<FamilyLinkTargetParam> targets,
            Action<int, string> progress)
        {
            if (app == null) throw new ArgumentNullException(nameof(app));

            var rows = new List<FamilyLinkAuditRow>();
            var targetMap = BuildTargetMap(targets);
            if (targetMap.Count == 0) return rows;

            var cleanedPaths = NormalizePaths(rvtPaths);
            var total = cleanedPaths.Count;
            if (total == 0) return rows;

            for (var i = 0; i < total; i++)
            {
                var rvtPath = cleanedPaths[i];
                Document doc = null;
                var fileName = SafeFileName(rvtPath);

                try
                {
                    ReportProgress(progress, total, i + 1, 0.02, $"문서 여는 중... {i + 1}/{total} {fileName}");

                    if (string.IsNullOrWhiteSpace(rvtPath)) throw new ArgumentException("RVT 경로가 비어 있습니다.");
                    if (!File.Exists(rvtPath)) throw new FileNotFoundException("RVT 파일을 찾을 수 없습니다.", rvtPath);

                    doc = OpenProjectDocument(app.Application, rvtPath);
                    if (doc == null) throw new InvalidOperationException("문서를 열 수 없습니다.");

                    var hostFamilies = new FilteredElementCollector(doc)
                        .OfClass(typeof(Family))
                        .Cast<Family>()
                        .Where(f => f != null && f.IsEditable && !f.IsInPlace)
                        .ToList();

                    var famTotal = hostFamilies.Count;
                    var rvtName = fileName;

                    if (famTotal == 0)
                    {
                        ReportProgress(progress, total, i + 1, 1.0, $"{rvtName}: 편집 가능한 패밀리가 없습니다.");
                        continue;
                    }

                    for (var fi = 0; fi < famTotal; fi++)
                    {
                        var fam = hostFamilies[fi];
                        var frac = 0.05 + 0.9 * SafeRatio(fi + 1, famTotal);
                        ReportProgress(progress, total, i + 1, frac, $"[{rvtName}] 패밀리 검사 중 ({fi + 1}/{famTotal})");
                        AuditFamilyAsHost(doc, fam, fileName, targetMap, rows);
                    }

                    ReportProgress(progress, total, i + 1, 1.0, $"완료: {rvtName}");
                }
                catch (Exception ex)
                {
                    rows.Add(new FamilyLinkAuditRow
                    {
                        FileName = fileName,
                        Issue = FamilyLinkAuditIssue.Error.ToString(),
                        Notes = $"Project open/scan error: {ex.Message}"
                    });
                }
                finally
                {
                    if (doc != null)
                    {
                        try { doc.Close(false); } catch { }
                    }
                }
            }

            return rows.Where(x => !string.Equals((x.Issue ?? "").Trim(), "OK", StringComparison.OrdinalIgnoreCase)).ToList();
        }

        public static List<FamilyLinkAuditRow> RunOnDocument(
            Document doc,
            string rvtPath,
            IList<FamilyLinkTargetParam> targets,
            Action<int, string> progress)
        {
            var rows = new List<FamilyLinkAuditRow>();
            if (doc == null) return rows;

            var targetMap = BuildTargetMap(targets);
            if (targetMap.Count == 0) return rows;

            var fileName = SafeFileName(rvtPath);
            try
            {
                var hostFamilies = new FilteredElementCollector(doc)
                    .OfClass(typeof(Family))
                    .Cast<Family>()
                    .Where(f => f != null && f.IsEditable && !f.IsInPlace)
                    .ToList();

                var famTotal = hostFamilies.Count;
                if (famTotal == 0)
                {
                    ReportProgress(progress, 1, 1, 1.0, $"{fileName}: 편집 가능한 패밀리가 없습니다.");
                    return rows;
                }

                for (var fi = 0; fi < famTotal; fi++)
                {
                    var fam = hostFamilies[fi];
                    var frac = 0.05 + 0.9 * SafeRatio(fi + 1, famTotal);
                    ReportProgress(progress, 1, 1, frac, $"[{fileName}] 패밀리 검사 중 ({fi + 1}/{famTotal})");
                    AuditFamilyAsHost(doc, fam, fileName, targetMap, rows);
                }

                ReportProgress(progress, 1, 1, 1.0, $"완료: {fileName}");
            }
            catch (Exception ex)
            {
                rows.Add(new FamilyLinkAuditRow
                {
                    FileName = fileName,
                    Issue = FamilyLinkAuditIssue.Error.ToString(),
                    Notes = $"Project scan error: {ex.Message}"
                });
            }

            return rows.Where(x => !string.Equals((x.Issue ?? "").Trim(), "OK", StringComparison.OrdinalIgnoreCase)).ToList();
        }

        private static void AuditFamilyAsHost(
            Document hostDoc,
            Family hostFamily,
            string fileName,
            Dictionary<string, FamilyLinkTargetParam> expectedByName,
            List<FamilyLinkAuditRow> rows)
        {
            Document famDoc = null;
            try
            {
                famDoc = hostDoc.EditFamily(hostFamily);
                if (famDoc == null || !famDoc.IsFamilyDocument) return;

                var nestedInstances = new FilteredElementCollector(famDoc)
                    .OfClass(typeof(FamilyInstance))
                    .WhereElementIsNotElementType()
                    .Cast<FamilyInstance>()
                    .Where(x => x != null && x.Symbol != null && x.Symbol.Family != null)
                    .ToList();

                if (nestedInstances.Count == 0) return;

                var repInstances = nestedInstances
                    .GroupBy(fi => fi.Symbol.Id.IntegerValue)
                    .Select(g => g.First())
                    .ToList();

                var hostCat = "";
                try
                {
                    if (hostFamily.FamilyCategory != null) hostCat = hostFamily.FamilyCategory.Name;
                }
                catch { }

                foreach (var fi in repInstances)
                {
                    var nestedFam = fi.Symbol.Family;
                    var nestedCat = "";
                    try
                    {
                        if (fi.Category != null) nestedCat = fi.Category.Name;
                    }
                    catch { }

                    var map = CollectParamMap(fi);

                    foreach (var kv in expectedByName)
                    {
                        var targetName = kv.Key;
                        var expected = kv.Value;

                        if (!map.TryGetValue(targetName, out var found) || found == null || found.Count == 0)
                        {
                            rows.Add(new FamilyLinkAuditRow
                            {
                                FileName = fileName,
                                HostFamilyName = hostFamily.Name,
                                HostFamilyCategory = hostCat,
                                NestedFamilyName = nestedFam.Name,
                                NestedTypeName = SafeStr(fi.Symbol.Name),
                                NestedCategory = nestedCat,
                                NestedParamName = "",
                                TargetParamName = targetName,
                                ExpectedGuid = expected.Guid.ToString("D"),
                                Issue = FamilyLinkAuditIssue.ParamNotFound.ToString(),
                                Notes = "네스티드(하위) 패밀리 인스턴스/타입에서 해당 이름의 파라미터를 찾지 못함"
                            });
                            continue;
                        }

                        foreach (var fp in found)
                        {
                            var p = fp.P;
                            if (p == null || p.Definition == null) continue;

                            var nestedGuidOk = TryGetParameterGuid(p, out var nestedGuid);
                            var nestedGuidStr = nestedGuidOk ? nestedGuid.ToString("D") : "";

                            var nestedIsSharedKnown = TryGetParameterIsShared(p, out var nestedIsShared);

                            FamilyParameter assoc;
                            try
                            {
                                assoc = famDoc.FamilyManager.GetAssociatedFamilyParameter(p);
                            }
                            catch (Exception ex)
                            {
                                rows.Add(new FamilyLinkAuditRow
                                {
                                    FileName = fileName,
                                    HostFamilyName = hostFamily.Name,
                                    HostFamilyCategory = hostCat,
                                    NestedFamilyName = nestedFam.Name,
                                    NestedTypeName = SafeStr(fi.Symbol.Name),
                                    NestedCategory = nestedCat,
                                    NestedParamName = SafeStr(p.Definition.Name),
                                    TargetParamName = targetName,
                                    ExpectedGuid = expected.Guid.ToString("D"),
                                    FoundScope = fp.Scope.ToString(),
                                    NestedParamGuid = nestedGuidStr,
                                    NestedParamDataType = SafeDefTypeToken(p.Definition),
                                    Issue = FamilyLinkAuditIssue.Error.ToString(),
                                    Notes = "GetAssociatedFamilyParameter 실패: " + ex.Message
                                });
                                continue;
                            }

                            var issue = FamilyLinkAuditIssue.OK;
                            var notes = "";

                            if (assoc == null)
                            {
                                issue = FamilyLinkAuditIssue.MissingAssociation;
                                notes = "호스트 패밀리 파라미터로 연동(Associate)되지 않음";
                            }
                            else
                            {
                                if (nestedGuidOk)
                                {
                                    if (nestedGuid != expected.Guid)
                                    {
                                        issue = FamilyLinkAuditIssue.GuidMismatch;
                                        notes = $"네스티드 파라미터 GUID 불일치 (Expected {expected.Guid:D}, Nested {nestedGuid:D})";
                                    }
                                }
                                else
                                {
                                    if (nestedIsSharedKnown)
                                    {
                                        notes = nestedIsShared
                                            ? "네스티드 파라미터 IsShared=True 이지만 GUID 추출 실패(특이 케이스)"
                                            : "네스티드 파라미터 IsShared=False (Shared 아님, 이름만 일치)";
                                    }
                                    else
                                    {
                                        notes = "네스티드 파라미터 Shared 여부 확인 실패(이름만 일치)";
                                    }
                                }

                                if (!assoc.IsShared)
                                {
                                    if (issue == FamilyLinkAuditIssue.OK) issue = FamilyLinkAuditIssue.HostParamNotShared;
                                    if (notes != "") notes += " / ";
                                    notes += "연결된 호스트 FamilyParameter가 Shared가 아님";
                                }

                                if (TryGetDefinitionGuid(assoc.Definition, out var hostGuid) && hostGuid != expected.Guid)
                                {
                                    if (issue == FamilyLinkAuditIssue.OK) issue = FamilyLinkAuditIssue.GuidMismatch;
                                    if (notes != "") notes += " / ";
                                    notes += $"호스트 파라미터 GUID 불일치 (Expected {expected.Guid:D}, Host {hostGuid:D})";
                                }
                            }

                            var row = new FamilyLinkAuditRow
                            {
                                FileName = fileName,
                                HostFamilyName = hostFamily.Name,
                                HostFamilyCategory = hostCat,
                                NestedFamilyName = nestedFam.Name,
                                NestedTypeName = SafeStr(fi.Symbol.Name),
                                NestedCategory = nestedCat,
                                NestedParamName = SafeStr(p.Definition.Name),
                                TargetParamName = targetName,
                                ExpectedGuid = expected.Guid.ToString("D"),
                                FoundScope = fp.Scope.ToString(),
                                NestedParamGuid = nestedGuidStr,
                                NestedParamDataType = SafeDefTypeToken(p.Definition),
                                Issue = issue.ToString(),
                                Notes = notes
                            };

                            if (assoc != null)
                            {
                                row.AssocHostParamName = SafeStr(assoc.Definition.Name);
                                row.HostParamIsShared = assoc.IsShared.ToString();

                                if (TryGetDefinitionGuid(assoc.Definition, out var hostGuid2))
                                {
                                    row.HostParamGuid = hostGuid2.ToString("D");
                                }
                            }

                            rows.Add(row);
                        }
                    }
                }
            }
            finally
            {
                if (famDoc != null)
                {
                    try { famDoc.Close(false); } catch { }
                }
            }
        }

        private static Dictionary<string, List<FoundParam>> CollectParamMap(FamilyInstance fi)
        {
            var map = new Dictionary<string, List<FoundParam>>(StringComparer.OrdinalIgnoreCase);

            try
            {
                foreach (Parameter p in fi.Parameters)
                {
                    if (p == null || p.Definition == null) continue;
                    var name = p.Definition.Name;
                    if (string.IsNullOrWhiteSpace(name)) continue;
                    if (!map.ContainsKey(name)) map[name] = new List<FoundParam>();
                    map[name].Add(new FoundParam { P = p, Scope = FoundScope.InstanceParam });
                }
            }
            catch { }

            try
            {
                if (fi.Symbol != null)
                {
                    foreach (Parameter p in fi.Symbol.Parameters)
                    {
                        if (p == null || p.Definition == null) continue;
                        var name = p.Definition.Name;
                        if (string.IsNullOrWhiteSpace(name)) continue;
                        if (!map.ContainsKey(name)) map[name] = new List<FoundParam>();
                        map[name].Add(new FoundParam { P = p, Scope = FoundScope.TypeParam });
                    }
                }
            }
            catch { }

            return map;
        }

        private static bool TryGetDefinitionGuid(Definition defn, out Guid guid)
        {
            guid = Guid.Empty;
            try
            {
                var ext = defn as ExternalDefinition;
                if (ext != null)
                {
                    guid = ext.GUID;
                    if (guid != Guid.Empty) return true;
                }
            }
            catch { }

            return false;
        }

        private static bool TryGetParameterIsShared(Parameter p, out bool isShared)
        {
            isShared = false;
            if (p == null) return false;

            try
            {
                var t = p.GetType();
                var prop = t.GetProperty("IsShared");
                if (prop == null) return false;
                var v = prop.GetValue(p, null);
                if (v is bool b)
                {
                    isShared = b;
                    return true;
                }
            }
            catch { }

            return false;
        }

        private static bool TryGetParameterGuid(Parameter p, out Guid guid)
        {
            guid = Guid.Empty;
            if (p == null) return false;

            try
            {
                var isSharedKnown = TryGetParameterIsShared(p, out var isShared);
                if (isSharedKnown && isShared)
                {
                    var t = p.GetType();
                    var propGuid = t.GetProperty("GUID") ?? t.GetProperty("Guid");
                    if (propGuid != null)
                    {
                        var v = propGuid.GetValue(p, null);
                        if (v is Guid g)
                        {
                            guid = g;
                            if (guid != Guid.Empty) return true;
                        }
                    }
                }
            }
            catch
            {
                // ignore reflection edge-cases and fallback
            }

            return TryGetDefinitionGuid(p.Definition, out guid);
        }

        private static string SafeDefTypeToken(Definition defn)
        {
            if (defn == null) return "";
            try
            {
#if REVIT2023
                var dt = defn.GetDataType();
                if (dt != null) return SafeStr(dt.TypeId);
                return "";
#else
                return SafeStr(defn.ParameterType.ToString());
#endif
            }
            catch
            {
                return "";
            }
        }

        private static Document OpenProjectDocument(Autodesk.Revit.ApplicationServices.Application app, string rvtPath)
        {
            if (string.IsNullOrWhiteSpace(rvtPath)) throw new ArgumentException("path is empty.");
            var mp = ModelPathUtils.ConvertUserVisiblePathToModelPath(rvtPath);

            var opts = new OpenOptions { Audit = false };
            try
            {
                opts.DetachFromCentralOption = DetachFromCentralOption.DetachAndPreserveWorksets;
            }
            catch { }

            try
            {
                var ws = new WorksetConfiguration(WorksetConfigurationOption.CloseAllWorksets);
                opts.SetOpenWorksetsConfiguration(ws);
            }
            catch { }

            try
            {
                return app.OpenDocumentFile(mp, opts);
            }
            catch
            {
                var opts2 = new OpenOptions { Audit = false };
                try
                {
                    var ws2 = new WorksetConfiguration(WorksetConfigurationOption.CloseAllWorksets);
                    opts2.SetOpenWorksetsConfiguration(ws2);
                }
                catch { }

                return app.OpenDocumentFile(mp, opts2);
            }
        }

        private static Dictionary<string, FamilyLinkTargetParam> BuildTargetMap(IList<FamilyLinkTargetParam> targets)
        {
            var map = new Dictionary<string, FamilyLinkTargetParam>(StringComparer.OrdinalIgnoreCase);
            if (targets == null) return map;

            foreach (var t in targets)
            {
                if (t == null) continue;
                if (string.IsNullOrWhiteSpace(t.Name)) continue;
                map[t.Name] = t;
            }

            return map;
        }

        private static List<string> NormalizePaths(IList<string> paths)
        {
            var list = new List<string>();
            var dedup = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            if (paths == null) return list;

            foreach (var p in paths)
            {
                if (string.IsNullOrWhiteSpace(p)) continue;
                string full;
                try
                {
                    full = Path.GetFullPath(p);
                }
                catch
                {
                    full = p;
                }

                if (dedup.Add(full)) list.Add(full);
            }

            return list;
        }

        private static string SafeFileName(string rvtPath)
        {
            if (string.IsNullOrWhiteSpace(rvtPath)) return "(Unknown)";
            try
            {
                return Path.GetFileName(rvtPath);
            }
            catch
            {
                return rvtPath;
            }
        }

        private static void ReportProgress(Action<int, string> progress, int total, int index, double fileProgress, string message)
        {
            if (progress == null) return;
            try
            {
                var ratio = 1.0;
                if (total > 0)
                {
                    var clamped = Math.Max(0.0, Math.Min(1.0, fileProgress));
                    ratio = (index - 1 + clamped) / total;
                }

                var pct = (int)Math.Max(0, Math.Min(100, Math.Round(ratio * 100)));
                progress(pct, message);
            }
            catch { }
        }

        private static double SafeRatio(int current, int total)
        {
            if (total <= 0) return 1.0;
            return Math.Max(0.0, Math.Min(1.0, (double)current / total));
        }

        private static string SafeStr(string s)
        {
            return s ?? "";
        }
    }
}
