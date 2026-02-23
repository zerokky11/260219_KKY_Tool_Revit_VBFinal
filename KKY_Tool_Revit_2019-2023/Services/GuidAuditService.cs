using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KKY_Tool_Revit.Infrastructure;

namespace KKY_Tool_Revit.Services
{
    public static class GuidAuditService
    {
        public sealed class RunResult
        {
            public int Mode { get; set; }
            public DataTable Project { get; set; }
            public DataTable FamilyDetail { get; set; }
            public DataTable FamilyIndex { get; set; }
            public string RunId { get; set; }
            public bool IncludeFamily { get; set; }
        }

        private sealed class TargetFile
        {
            public string Path { get; set; } = string.Empty;
            public string Name { get; set; } = string.Empty;
        }

        private sealed class FamilyAuditPack
        {
            public DataTable Summary { get; set; }
            public DataTable Detail { get; set; }
            public DataTable Index { get; set; }
        }

        public static RunResult Run(
            UIApplication app,
            int mode,
            IEnumerable<string> rvtPaths,
            Action<double, string> progress,
            Action<string> warn = null,
            bool includeFamily = false,
            bool includeAnnotation = false)
        {
            if (app == null) throw new ArgumentNullException(nameof(app));

            var defMap = SharedParamReader.ReadSharedParamNameGuidMap(app.Application);
            if (defMap == null || defMap.Count == 0)
            {
                throw new InvalidOperationException("공유 파라미터 파일이 설정되어 있지 않거나 읽을 수 없습니다. (Revit 옵션에서 Shared Parameter 파일 경로 확인)");
            }

            var targets = BuildTargets(app, rvtPaths);
            if (targets.Count == 0)
            {
                throw new InvalidOperationException("검토할 RVT 파일이 없습니다.");
            }

            var runId = Guid.NewGuid().ToString("N");
            var total = targets.Count;
            DataTable projectTable = null;
            DataTable familyDetail = null;
            DataTable famIndex = null;

            for (var i = 0; i < total; i++)
            {
                var target = targets[i];
                var openedByMe = false;
                Document doc = null;
                var openError = string.Empty;

                try
                {
                    ReportProgress(progress, total, i + 1, 0.02, $"문서 여는 중... {i + 1}/{total} {target.Name}");
                    doc = ResolveOrOpenDocument(app, app.ActiveUIDocument?.Document, target.Path, ref openedByMe, ref openError);

                    if (doc == null)
                    {
                        var failProj = Auditors.MakeFailureSummaryTable(1);
                        var note = BuildOpenFailNotes(openError, target.Path);
                        var shortReason = ShortenReason(note);
                        if (warn != null && !string.IsNullOrWhiteSpace(note))
                        {
                            warn(note);
                        }
                        ReportProgress(progress, total, i + 1, 0.08, $"문서 열기 실패: {target.Name} - {shortReason}");
                        Auditors.AddOpenFailRow(failProj, target.Name, target.Path, "Project", "OPEN_FAIL", note);
                        projectTable = MergeTable(projectTable, failProj);

                        if (includeFamily)
                        {
                            var failFam = Auditors.MakeFailureSummaryTable(2);
                            Auditors.AddOpenFailRow(failFam, target.Name, target.Path, "Family", "OPEN_FAIL", note);
                            familyDetail = MergeTable(familyDetail, failFam);
                        }
                        continue;
                    }

                    var rvtName = GetRvtName(doc, target.Path);
                    var captureIndex = i;
                    var captureName = rvtName;

                    var proj = Auditors.RunProjectParameterAudit(doc, defMap, rvtName, target.Path,
                        (cur, tot) =>
                        {
                            var frac = 0.1 + 0.8 * SafeRatio(cur, tot);
                            ReportProgress(progress, total, captureIndex + 1, frac, $"[{captureName}] 프로젝트 파라미터 ({cur}/{tot})");
                        });
                    projectTable = MergeTable(projectTable, proj);

                    if (includeFamily)
                    {
                        var famPack = Auditors.RunFamilyAudit(doc, defMap, rvtName, target.Path,
                            (cur, tot, famName) =>
                            {
                                var frac = 0.1 + 0.8 * SafeRatio(cur, tot);
                                ReportProgress(progress, total, captureIndex + 1, frac, $"[{captureName}] 패밀리 처리 중 ({cur}/{tot}) {famName}");
                            }, includeAnnotation);

                        familyDetail = MergeTable(familyDetail, famPack?.Detail);
                        famIndex = MergeTable(famIndex, famPack?.Index);
                    }

                    ReportProgress(progress, total, i + 1, 1.0, $"완료: {target.Name}");
                }
                catch (Exception ex)
                {
                    var note = BuildExceptionNotes(ex, target.Path);
                    if (warn != null && !string.IsNullOrWhiteSpace(note))
                    {
                        warn(note);
                    }

                    var fail = Auditors.MakeFailureSummaryTable(1);
                    Auditors.AddOpenFailRow(fail, target.Name, target.Path, "Project", "ERROR", note);
                    projectTable = MergeTable(projectTable, fail);

                    if (includeFamily)
                    {
                        var failFam = Auditors.MakeFailureSummaryTable(2);
                        Auditors.AddOpenFailRow(failFam, target.Name, target.Path, "Family", "ERROR", note);
                        familyDetail = MergeTable(familyDetail, failFam);
                    }
                }
                finally
                {
                    if (openedByMe && doc != null)
                    {
                        try { doc.Close(false); } catch { }
                    }
                }
            }

            ResultTableFilter.KeepOnlyIssues("guid", projectTable);
            if (includeFamily)
            {
                ResultTableFilter.KeepOnlyIssues("guid", familyDetail);

                var famSet = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                if (familyDetail != null && familyDetail.Columns.Contains("FamilyName"))
                {
                    foreach (DataRow r in familyDetail.Rows)
                    {
                        var fn = Convert.ToString(r["FamilyName"]).Trim();
                        if (!string.IsNullOrWhiteSpace(fn)) famSet.Add(fn);
                    }
                }
                ResultTableFilter.KeepOnlyByNameSet(famIndex, "FamilyName", famSet);
            }

            return new RunResult
            {
                Mode = mode,
                Project = projectTable ?? Auditors.MakeFailureSummaryTable(1),
                FamilyDetail = includeFamily ? familyDetail : null,
                FamilyIndex = includeFamily ? famIndex : null,
                RunId = runId,
                IncludeFamily = includeFamily
            };
        }

        public static string Export(DataTable table, string sheetName, string excelMode = "fast", string progressChannel = null)
        {
            if (table == null) return string.Empty;
            var doAutoFit = false;
            try
            {
                if (string.Equals(excelMode, "normal", StringComparison.OrdinalIgnoreCase) && table.Rows.Count <= 30000)
                {
                    doAutoFit = true;
                }
            }
            catch { doAutoFit = false; }

            ResultTableFilter.KeepOnlyIssues("guid", table);
            ExcelCore.EnsureMessageRow(table, "오류가 없습니다.");

            using (var sfd = new SaveFileDialog())
            {
                sfd.Filter = "Excel Workbook (*.xlsx)|*.xlsx";
                sfd.FileName = $"{sheetName}_{DateTime.Now:yyyyMMdd_HHmm}.xlsx";
                if (sfd.ShowDialog() != DialogResult.OK) return string.Empty;
                ExcelCore.SaveXlsx(sfd.FileName, sheetName, table, doAutoFit, sheetKey: sheetName, progressKey: progressChannel);
                return sfd.FileName;
            }
        }

        public static string ExportMulti(IList<KeyValuePair<string, DataTable>> sheets, string excelMode = "fast", string progressChannel = null)
        {
            if (sheets == null || sheets.Count == 0) return string.Empty;
            var doAutoFit = false;
            try
            {
                if (string.Equals(excelMode, "normal", StringComparison.OrdinalIgnoreCase)) doAutoFit = true;
            }
            catch { doAutoFit = false; }

            foreach (var kv in sheets)
            {
                ResultTableFilter.KeepOnlyIssues("guid", kv.Value);
                ExcelCore.EnsureMessageRow(kv.Value, "오류가 없습니다.");
            }

            using (var sfd = new SaveFileDialog())
            {
                sfd.Filter = "Excel Workbook (*.xlsx)|*.xlsx";
                sfd.FileName = $"GuidAudit_{DateTime.Now:yyyyMMdd_HHmm}.xlsx";
                if (sfd.ShowDialog() != DialogResult.OK) return string.Empty;
                ExcelCore.SaveXlsxMulti(sfd.FileName, sheets, doAutoFit, progressChannel);
                return sfd.FileName;
            }
        }

        public static DataTable PrepareExportTable(DataTable source, int mode)
        {
            var baseTable = source ?? Auditors.MakeFailureSummaryTable(mode);
            var exportTable = baseTable.Clone();

            if (source != null)
            {
                foreach (DataRow r in source.Rows) exportTable.ImportRow(r);
            }

            ResultTableFilter.KeepOnlyIssues("guid", exportTable);
            if (exportTable.Columns.Contains("RvtPath")) exportTable.Columns.Remove("RvtPath");
            if (exportTable.Columns.Contains("BoundCategories")) exportTable.Columns["BoundCategories"].SetOrdinal(exportTable.Columns.Count - 1);
            ExcelCore.EnsureMessageRow(exportTable, "오류가 없습니다.");
            return exportTable;
        }

        private static List<TargetFile> BuildTargets(UIApplication app, IEnumerable<string> rvtPaths)
        {
            var list = new List<TargetFile>();
            var dedup = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var requested = rvtPaths ?? Enumerable.Empty<string>();

            foreach (var p in requested)
            {
                if (string.IsNullOrWhiteSpace(p)) continue;
                var full = p;
                try
                {
                    full = Path.IsPathRooted(p) ? Path.GetFullPath(p) : p.Trim();
                }
                catch
                {
                    full = p;
                }

                if (dedup.Add(full)) list.Add(new TargetFile { Path = full, Name = SafeFileName(full) });
            }

            if (list.Count == 0)
            {
                var ap = string.Empty;
                try { ap = app.ActiveUIDocument?.Document?.PathName; } catch { ap = string.Empty; }
                list.Add(new TargetFile { Path = ap, Name = SafeFileName(ap) });
            }

            return list;
        }

        private static string SafeFileName(string p)
        {
            if (string.IsNullOrWhiteSpace(p)) return "(Active/Unsaved)";
            try { return Path.GetFileName(p); } catch { return p; }
        }

        private static string GetRvtName(Document doc, string path)
        {
            if (!string.IsNullOrWhiteSpace(path))
            {
                try { return Path.GetFileName(path); } catch { }
            }

            try { return doc.Title; } catch { return "(Doc)"; }
        }

        private static DataTable MergeTable(DataTable master, DataTable part)
        {
            if (part == null) return master;
            if (master == null) master = part.Clone();
            foreach (DataRow r in part.Rows) master.ImportRow(r);
            return master;
        }

        private static double SafeRatio(int cur, int tot)
        {
            if (tot <= 0) return 0;
            return Math.Max(0, Math.Min(1.0, (double)cur / tot));
        }

        private static string NormalizeName(string s)
        {
            if (s == null) return string.Empty;
            var value = s.Replace('\u00A0', ' ').Trim();
            if (value.Length == 0) return string.Empty;
            try { value = Regex.Replace(value, "\\s+", " "); } catch { }
            return value;
        }

        private static void ReportProgress(Action<double, string> cb, int totalFiles, int fileIndex, double docProgress, string text)
        {
            if (cb == null) return;
            var safeTotal = Math.Max(1, totalFiles);
            var idx = Math.Max(0, fileIndex - 1);
            var ratio = (idx + Math.Max(0.0, Math.Min(1.0, docProgress))) / safeTotal;
            var pct = Math.Max(0, Math.Min(100, Math.Round(ratio * 1000.0) / 10.0));
            cb(pct, text);
        }

        private static string BuildOpenFailNotes(string reason, string inputPath)
        {
            var trimmed = (reason ?? string.Empty).Trim();
            var hasPathInReason = false;
            try
            {
                hasPathInReason = !string.IsNullOrWhiteSpace(inputPath) &&
                                  trimmed.IndexOf(inputPath, StringComparison.OrdinalIgnoreCase) >= 0;
            }
            catch { hasPathInReason = false; }

            var pathPart = string.IsNullOrWhiteSpace(inputPath) || hasPathInReason ? string.Empty : $" [Path: {inputPath}]";
            if (string.IsNullOrWhiteSpace(trimmed)) return $"문서 열기 실패{pathPart}";
            return $"{trimmed}{pathPart}";
        }

        private static string BuildExceptionNotes(Exception ex, string inputPath)
        {
            if (ex == null) return BuildOpenFailNotes(string.Empty, inputPath);
            string hrPart;
            try { hrPart = $" (0x{ex.HResult:X8})"; } catch { hrPart = string.Empty; }
            return BuildOpenFailNotes($"{ex.Message}{hrPart}", inputPath);
        }

        private static string ShortenReason(string reason)
        {
            if (string.IsNullOrWhiteSpace(reason)) return string.Empty;
            var firstLine = reason.Replace("\r", " ").Replace("\n", " ").Trim();
            return firstLine.Length > 120 ? firstLine.Substring(0, 117) + "..." : firstLine;
        }

        private static Document ResolveOrOpenDocument(UIApplication uiApp, Document activeDoc, string path, ref bool openedByMe, ref string failureReason)
        {
            openedByMe = false;
            failureReason = string.Empty;

            var requested = (path ?? string.Empty).Trim();
            bool isRooted;
            try { isRooted = Path.IsPathRooted(requested); } catch { isRooted = false; }
            var allowNameMatch = !isRooted && requested.IndexOf(':') == -1 && requested.IndexOf('\\') == -1;

            if (string.IsNullOrWhiteSpace(requested)) return activeDoc;
            if (IsMatchingDoc(activeDoc, requested, allowNameMatch)) return activeDoc;

            var opened = FindOpenDocument(uiApp, requested, allowNameMatch);
            if (opened != null) return opened;

            if (allowNameMatch || !isRooted)
            {
                failureReason = $"Invalid path: {requested}";
                return null;
            }

            try
            {
                if (Path.IsPathRooted(requested) && !File.Exists(requested))
                {
                    failureReason = $"File not found: {requested}";
                    return null;
                }
            }
            catch (Exception ex)
            {
                failureReason = BuildExceptionNotes(ex, requested);
                return null;
            }

            ModelPath mp = null;
            try { mp = ModelPathUtils.ConvertUserVisiblePathToModelPath(requested); }
            catch (Exception ex)
            {
                failureReason = BuildExceptionNotes(ex, requested);
                mp = null;
            }

            if (mp == null)
            {
                if (string.IsNullOrWhiteSpace(failureReason)) failureReason = $"경로 변환 실패 [Path: {requested}]";
                return null;
            }

            var preferDetach = false;
            try
            {
                var bfi = BasicFileInfo.Extract(requested);
                preferDetach = bfi == null || bfi.IsWorkshared;
            }
            catch { preferDetach = true; }

            var attempts = new List<OpenOptions>();
            if (preferDetach) attempts.Add(CreateDetachOptions());
            attempts.Add(new OpenOptions());

            foreach (var opt in attempts)
            {
                try
                {
                    var d = uiApp.Application.OpenDocumentFile(mp, opt);
                    openedByMe = true;
                    failureReason = string.Empty;
                    return d;
                }
                catch (Exception ex)
                {
                    failureReason = BuildExceptionNotes(ex, requested);
                }
            }

            openedByMe = false;
            return null;
        }

        private static OpenOptions CreateDetachOptions()
        {
            var opt = new OpenOptions();
            try
            {
                opt.DetachFromCentralOption = DetachFromCentralOption.DetachAndPreserveWorksets;
                var wc = new WorksetConfiguration(WorksetConfigurationOption.CloseAllWorksets);
                opt.SetOpenWorksetsConfiguration(wc);
            }
            catch { }
            return opt;
        }

        private static Document FindOpenDocument(UIApplication uiApp, string requested, bool allowNameMatch)
        {
            if (uiApp == null) return null;
            try
            {
                foreach (Document d in uiApp.Application.Documents)
                {
                    if (IsMatchingDoc(d, requested, allowNameMatch)) return d;
                }
            }
            catch { }
            return null;
        }

        private static bool IsMatchingDoc(Document doc, string requested, bool allowNameMatch)
        {
            if (doc == null) return false;

            string dp;
            try { dp = doc.PathName; } catch { dp = string.Empty; }
            if (!string.IsNullOrWhiteSpace(dp) && string.Equals(dp, requested, StringComparison.OrdinalIgnoreCase)) return true;

            if (!allowNameMatch) return false;

            string fileOnly;
            try { fileOnly = Path.GetFileName(dp); } catch { fileOnly = string.Empty; }
            if (!string.IsNullOrWhiteSpace(fileOnly) && string.Equals(fileOnly, requested, StringComparison.OrdinalIgnoreCase)) return true;

            string title;
            try { title = doc.Title; } catch { title = string.Empty; }
            return !string.IsNullOrWhiteSpace(title) && string.Equals(title, requested, StringComparison.OrdinalIgnoreCase);
        }

        private static class SharedParamReader
        {
            public static Dictionary<string, List<Guid>> ReadSharedParamNameGuidMap(Autodesk.Revit.ApplicationServices.Application app)
            {
                DefinitionFile defFile = null;
                try
                {
                    defFile = app.OpenSharedParameterFile();
                }
                catch
                {
                    defFile = null;
                }

                if (defFile == null) return null;

                var map = new Dictionary<string, List<Guid>>(StringComparer.OrdinalIgnoreCase);
                foreach (DefinitionGroup grp in defFile.Groups)
                {
                    foreach (Definition d in grp.Definitions)
                    {
                        if (!TryGetDefinitionGuid(d, out var g)) continue;
                        var name = NormalizeName(d.Name);
                        if (!map.ContainsKey(name)) map[name] = new List<Guid>();
                        map[name].Add(g);
                    }
                }

                return map;
            }

            private static bool TryGetDefinitionGuid(Definition d, out Guid g)
            {
                g = Guid.Empty;
                if (d == null) return false;

                var t = d.GetType();
                var p = t.GetProperty("GUID", BindingFlags.Public | BindingFlags.Instance);
                if (p == null) return false;

                var v = p.GetValue(d, null);
                if (v == null) return false;

                if (v is Guid gv)
                {
                    g = gv;
                    return g != Guid.Empty;
                }

                return false;
            }
        }

        private static class Auditors
        {
            public static DataTable MakeFailureSummaryTable(int mode)
            {
                if (mode == 1)
                {
                    var dt = new DataTable("ProjectParams");
                    dt.Columns.Add("RvtName", typeof(string));
                    dt.Columns.Add("RvtPath", typeof(string));
                    dt.Columns.Add("ParamName", typeof(string));
                    dt.Columns.Add("ParamKind", typeof(string));
                    dt.Columns.Add("ParamGroup", typeof(string));
                    dt.Columns.Add("BoundCategories", typeof(string));
                    dt.Columns.Add("RvtGuid", typeof(string));
                    dt.Columns.Add("FileGuid", typeof(string));
                    dt.Columns.Add("Result", typeof(string));
                    dt.Columns.Add("Notes", typeof(string));
                    return dt;
                }

                var dt2 = new DataTable("FamilyParams");
                dt2.Columns.Add("RvtName", typeof(string));
                dt2.Columns.Add("RvtPath", typeof(string));
                dt2.Columns.Add("FamilyName", typeof(string));
                dt2.Columns.Add("FamilyCategory", typeof(string));
                dt2.Columns.Add("ParamName", typeof(string));
                dt2.Columns.Add("IsShared", typeof(string));
                dt2.Columns.Add("FamilyGuid", typeof(string));
                dt2.Columns.Add("FileGuid", typeof(string));
                dt2.Columns.Add("Result", typeof(string));
                dt2.Columns.Add("Notes", typeof(string));
                return dt2;
            }

            public static void AddOpenFailRow(DataTable dt, string rvtName, string rvtPath, string scope, string result, string notes)
            {
                var r = dt.NewRow();
                if (dt.Columns.Contains("RvtName")) r["RvtName"] = rvtName ?? string.Empty;
                if (dt.Columns.Contains("RvtPath")) r["RvtPath"] = rvtPath ?? string.Empty;
                if (dt.Columns.Contains("FamilyName")) r["FamilyName"] = string.Empty;
                if (dt.Columns.Contains("FamilyCategory")) r["FamilyCategory"] = string.Empty;
                if (dt.Columns.Contains("ParamName")) r["ParamName"] = string.Empty;
                if (dt.Columns.Contains("ParamKind")) r["ParamKind"] = string.Empty;
                if (dt.Columns.Contains("ParamGroup")) r["ParamGroup"] = string.Empty;
                if (dt.Columns.Contains("BoundCategories")) r["BoundCategories"] = string.Empty;
                if (dt.Columns.Contains("RvtGuid")) r["RvtGuid"] = string.Empty;
                if (dt.Columns.Contains("IsShared")) r["IsShared"] = string.Empty;
                if (dt.Columns.Contains("FamilyGuid")) r["FamilyGuid"] = string.Empty;
                if (dt.Columns.Contains("FileGuid")) r["FileGuid"] = string.Empty;
                if (dt.Columns.Contains("Result")) r["Result"] = result;
                if (dt.Columns.Contains("Notes")) r["Notes"] = notes;
                dt.Rows.Add(r);
            }

            public static DataTable RunProjectParameterAudit(Document doc,
                Dictionary<string, List<Guid>> fileMap,
                string rvtName,
                string rvtPath,
                Action<int, int> progress = null)
            {
                var dt = MakeFailureSummaryTable(1);
                var allowedCategoryNames = BuildAllowedCategoryNameSet(doc);

                var speByName = new Dictionary<string, List<Guid>>(StringComparer.OrdinalIgnoreCase);
                try
                {
                    foreach (SharedParameterElement spe in new FilteredElementCollector(doc).OfClass(typeof(SharedParameterElement)).Cast<SharedParameterElement>())
                    {
                        var key = NormalizeName(SafeParamElementName(spe));
                        Guid g;
                        try { g = spe.GuidValue; } catch { g = Guid.Empty; }
                        if (g == Guid.Empty) continue;
                        if (!speByName.ContainsKey(key)) speByName[key] = new List<Guid>();
                        speByName[key].Add(g);
                    }
                }
                catch { }

                var bindings = doc.ParameterBindings;
                var iter = bindings.ForwardIterator();
                iter.Reset();

                var idx = 0;
                var total = 0;
                try { while (iter.MoveNext()) total++; } catch { total = 0; }
                try { iter.Reset(); } catch { }

                while (true)
                {
                    bool moved;
                    try { moved = iter.MoveNext(); } catch { break; }
                    if (!moved) break;

                    idx++;
                    progress?.Invoke(idx, Math.Max(1, total));

                    Definition def;
                    ElementBinding binding;
                    try
                    {
                        def = iter.Key;
                        binding = iter.Current as ElementBinding;
                    }
                    catch
                    {
                        def = null;
                        binding = null;
                    }

                    if (def == null) continue;

                    string name;
                    try { name = def.Name; } catch { name = string.Empty; }
                    var normName = NormalizeName(name);

                    var kind = "Project";
                    var projGuid = string.Empty;
                    var fileGuid = string.Empty;
                    var result = string.Empty;
                    var notes = string.Empty;

                    var isShared = def is ExternalDefinition;
                    var docGuid = Guid.Empty;
                    List<Guid> docGuids = null;

                    if (isShared)
                    {
                        kind = "Shared";
                        try { docGuid = ((ExternalDefinition)def).GUID; } catch { docGuid = Guid.Empty; }
                        if (docGuid != Guid.Empty) docGuids = new List<Guid> { docGuid };
                    }
                    else if (speByName.TryGetValue(normName, out var list))
                    {
                        isShared = true;
                        kind = "Shared";
                        docGuids = new List<Guid>(list);
                        docGuid = docGuids.FirstOrDefault();
                    }

                    if (isShared)
                    {
                        projGuid = docGuid == Guid.Empty ? string.Empty : docGuid.ToString();
                        if (fileMap != null && fileMap.TryGetValue(normName, out var fileGuids))
                        {
                            fileGuid = string.Join("; ", fileGuids.Select(x => x.ToString()).Distinct().ToArray());
                            if (docGuids == null) docGuids = new List<Guid>();
                            var hit = fileGuids.Any(g => docGuids.Any(x => x == g));
                            result = hit ? (fileGuids.Count > 1 ? "OK(MULTI_IN_FILE)" : "OK") : "MISMATCH";
                            if (result == "MISMATCH") notes = "RVT의 GUID와 Shared Parameter 파일 GUID 불일치";
                        }
                        else
                        {
                            result = "NOT_FOUND_IN_FILE";
                            notes = "Shared Parameter 파일에서 동일 이름을 찾지 못함";
                        }

                        if (result == "OK" || result == "OK(MULTI_IN_FILE)" || result == "MISMATCH")
                        {
                            if (fileMap != null && fileMap.TryGetValue(normName, out var fgs) && fgs != null && fgs.Count > 1)
                                notes = AppendNote(notes, "파일 내 동일 이름 GUID 여러 개");
                            if (docGuids != null && docGuids.Count > 1)
                                notes = AppendNote(notes, "문서 내 동일 이름 GUID 여러 개");
                        }
                    }
                    else
                    {
                        result = "PROJECT_PARAM";
                    }

                    var r = dt.NewRow();
                    r["RvtName"] = rvtName ?? string.Empty;
                    r["RvtPath"] = rvtPath ?? string.Empty;
                    r["ParamName"] = name ?? string.Empty;
                    r["ParamKind"] = kind;
                    r["ParamGroup"] = SafeParameterGroupName(def);
                    r["BoundCategories"] = FormatBoundCategories(binding, allowedCategoryNames);
                    r["RvtGuid"] = projGuid;
                    r["FileGuid"] = fileGuid;
                    r["Result"] = result;
                    r["Notes"] = notes;
                    dt.Rows.Add(r);
                }

                if (dt.Columns.Contains("BoundCategories")) dt.Columns["BoundCategories"].SetOrdinal(dt.Columns.Count - 1);
                return dt;
            }

            public static FamilyAuditPack RunFamilyAudit(Document doc,
                Dictionary<string, List<Guid>> fileMap,
                string rvtName,
                string rvtPath,
                Action<int, int, string> progress = null,
                bool includeAnnotation = false)
            {
                var pack = new FamilyAuditPack();

                var dtDet = new DataTable("FamilyParamDetail");
                dtDet.Columns.Add("RvtName", typeof(string));
                dtDet.Columns.Add("RvtPath", typeof(string));
                dtDet.Columns.Add("FamilyName", typeof(string));
                dtDet.Columns.Add("FamilyCategory", typeof(string));
                dtDet.Columns.Add("ParamName", typeof(string));
                dtDet.Columns.Add("ParamKind", typeof(string));
                dtDet.Columns.Add("IsShared", typeof(string));
                dtDet.Columns.Add("FamilyGuid", typeof(string));
                dtDet.Columns.Add("FileGuid", typeof(string));
                dtDet.Columns.Add("Result", typeof(string));
                dtDet.Columns.Add("Notes", typeof(string));

                var dtIdx = new DataTable("FamilyIndex");
                dtIdx.Columns.Add("RvtName", typeof(string));
                dtIdx.Columns.Add("RvtPath", typeof(string));
                dtIdx.Columns.Add("FamilyName", typeof(string));
                dtIdx.Columns.Add("FamilyCategory", typeof(string));
                dtIdx.Columns.Add("TotalParamCount", typeof(int));
                dtIdx.Columns.Add("SharedParamCount", typeof(int));

                var fams = new FilteredElementCollector(doc)
                    .OfClass(typeof(Family))
                    .Cast<Family>()
                    .OrderBy(x => x.Name, StringComparer.OrdinalIgnoreCase)
                    .ToList();

                var total = Math.Max(1, fams.Count);
                var idx = 0;

                foreach (var fam in fams)
                {
                    idx++;
                    progress?.Invoke(idx, total, fam.Name);

                    var famName = fam.Name;
                    string famCat;
                    try { famCat = fam.FamilyCategory?.Name ?? string.Empty; } catch { famCat = string.Empty; }

                    try
                    {
                        if (fam.FamilyCategory != null)
                        {
                            CategoryType catType;
                            try { catType = fam.FamilyCategory.CategoryType; } catch { catType = (CategoryType)(-1); }
                            if (catType == CategoryType.Annotation && !includeAnnotation) continue;
                        }
                    }
                    catch { }

                    var skip = false;
                    try { if (fam.IsInPlace) skip = true; } catch { skip = false; }
                    if (skip) continue;

                    try
                    {
                        var p = fam.GetType().GetProperty("IsEditable", BindingFlags.Public | BindingFlags.Instance);
                        if (p != null)
                        {
                            var v = p.GetValue(fam, null);
                            if (v is bool b && !b) continue;
                        }
                    }
                    catch { }

                    Document famDoc = null;
                    try
                    {
                        var isInPlace = false;
                        try { isInPlace = fam.IsInPlace; } catch { }
                        if (isInPlace) continue;

                        try { famDoc = doc.EditFamily(fam); }
                        catch (InvalidOperationException)
                        {
                            famDoc = null;
                            continue;
                        }

                        if (famDoc == null || !famDoc.IsFamilyDocument) continue;

                        var fm = famDoc.FamilyManager;
                        if (fm == null)
                        {
                            AddDetailRow(dtDet, rvtName, rvtPath, famName, famCat, "", "N/A", "", "", "", "OPEN_FAIL", "FamilyManager 없음");
                            continue;
                        }

                        var totalParamCount = 0;
                        var sharedCount = 0;

                        foreach (FamilyParameter fp in fm.Parameters)
                        {
                            if (fp == null) continue;
                            totalParamCount++;

                            string pName;
                            try { pName = fp.Definition.Name; } catch { pName = string.Empty; }
                            var normParamName = NormalizeName(pName);

                            var paramKind = GetFamilyParamKind(fp);
                            var isSharedBool = string.Equals(paramKind, "Shared", StringComparison.OrdinalIgnoreCase);
                            if (isSharedBool) sharedCount++;

                            var famGuid = string.Empty;
                            var fileGuid = string.Empty;
                            var res = string.Empty;
                            var notes = string.Empty;

                            if (isSharedBool)
                            {
                                if (TryGetFamilyParameterGuid(fp, out var gFam))
                                {
                                    famGuid = gFam.ToString();
                                    if (fileMap != null && fileMap.TryGetValue(normParamName, out var fileGuids))
                                    {
                                        fileGuid = string.Join("; ", fileGuids.Select(x => x.ToString()).Distinct().ToArray());
                                        res = fileGuids.Any(x => x == gFam) ? (fileGuids.Count > 1 ? "OK(MULTI_IN_FILE)" : "OK") : "MISMATCH";
                                    }
                                    else
                                    {
                                        res = "NOT_FOUND_IN_FILE";
                                    }
                                }
                                else
                                {
                                    res = "GUID_FAIL";
                                    notes = "FamilyParameter GUID 추출 실패";
                                }
                            }
                            else if (string.Equals(paramKind, "BuiltIn", StringComparison.OrdinalIgnoreCase)) res = "BUILTIN";
                            else res = "FAMILY_PARAM";

                            if (isSharedBool)
                            {
                                if (res == "NOT_FOUND_IN_FILE") notes = "Shared Parameter 파일에서 동일 이름을 찾지 못함";
                                else if (res == "MISMATCH") notes = "RVT의 GUID와 Shared Parameter 파일 GUID 불일치";

                                if ((res == "OK" || res == "OK(MULTI_IN_FILE)" || res == "MISMATCH") &&
                                    fileMap != null && fileMap.TryGetValue(normParamName, out var fgs) && fgs != null && fgs.Count > 1)
                                {
                                    notes = AppendNote(notes, "파일 내 동일 이름 GUID 여러 개");
                                }
                            }

                            AddDetailRow(dtDet, rvtName, rvtPath, famName, famCat, pName, paramKind,
                                isSharedBool ? "Y" : "N", famGuid, fileGuid, res, notes);
                        }

                        var rIdx = dtIdx.NewRow();
                        rIdx["RvtName"] = rvtName ?? string.Empty;
                        rIdx["RvtPath"] = rvtPath ?? string.Empty;
                        rIdx["FamilyName"] = famName ?? string.Empty;
                        rIdx["FamilyCategory"] = famCat ?? string.Empty;
                        rIdx["TotalParamCount"] = totalParamCount;
                        rIdx["SharedParamCount"] = sharedCount;
                        dtIdx.Rows.Add(rIdx);
                    }
                    catch (Exception ex)
                    {
                        AddDetailRow(dtDet, rvtName, rvtPath, famName, famCat, "", "N/A", "", "", "", "OPEN_FAIL", ex.Message);
                    }
                    finally
                    {
                        if (famDoc != null)
                        {
                            try { famDoc.Close(false); } catch { }
                        }
                    }
                }

                pack.Summary = null;
                pack.Detail = dtDet;
                pack.Index = dtIdx;
                return pack;
            }

            private static string AppendNote(string existing, string note)
            {
                if (string.IsNullOrWhiteSpace(existing)) return note;
                if (string.IsNullOrWhiteSpace(note)) return existing;
                return existing + "; " + note;
            }

            private static string SafeParameterGroupName(Definition def)
            {
                try
                {
                    var idef = def as InternalDefinition;
                    if (idef != null)
                    {
                        var pg = idef.ParameterGroup;
                        try
                        {
                            var label = LabelUtils.GetLabelFor(pg);
                            if (!string.IsNullOrWhiteSpace(label)) return label;
                        }
                        catch { }
                        return pg.ToString();
                    }
                }
                catch { }
                return string.Empty;
            }

            private static string FormatBoundCategories(ElementBinding binding, HashSet<string> allowedCategoryNames)
            {
                if (binding?.Categories == null) return string.Empty;
                if (allowedCategoryNames == null || allowedCategoryNames.Count == 0) return string.Empty;

                var topLevelNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                var subByTop = new Dictionary<string, HashSet<string>>(StringComparer.OrdinalIgnoreCase);

                foreach (Category cat in binding.Categories)
                {
                    if (cat == null) continue;
                    var currentName = SafeCategoryName(cat);
                    if (string.IsNullOrWhiteSpace(currentName)) continue;

                    Category parent;
                    try { parent = cat.Parent; } catch { parent = null; }

                    if (parent == null)
                    {
                        if (allowedCategoryNames.Contains(currentName)) topLevelNames.Add(currentName);
                        continue;
                    }

                    var parentName = SafeCategoryName(parent);
                    if (string.IsNullOrWhiteSpace(parentName)) continue;
                    if (!allowedCategoryNames.Contains(parentName)) continue;

                    topLevelNames.Add(parentName);
                    if (!subByTop.ContainsKey(parentName)) subByTop[parentName] = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                    subByTop[parentName].Add(currentName);
                }

                var labels = new List<string>();
                foreach (var top in topLevelNames.OrderBy(x => x, StringComparer.OrdinalIgnoreCase))
                {
                    labels.Add($"[{top}]");
                    if (subByTop.TryGetValue(top, out var subs) && subs != null)
                    {
                        foreach (var subName in subs.OrderBy(x => x, StringComparer.OrdinalIgnoreCase))
                        {
                            labels.Add($"[{top}: {subName}]");
                        }
                    }
                }

                return string.Join(",", labels.ToArray());
            }

            private static string SafeCategoryName(Category cat)
            {
                if (cat == null) return string.Empty;
                try
                {
                    var name = cat.Name ?? string.Empty;
                    if (string.IsNullOrWhiteSpace(name)) return string.Empty;
                    return name.Trim();
                }
                catch { return string.Empty; }
            }

            private static HashSet<string> BuildAllowedCategoryNameSet(Document doc)
            {
                var result = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                if (doc == null) return result;

                Categories cats;
                try { cats = doc.Settings.Categories; } catch { cats = null; }
                if (cats == null) return result;

                foreach (Category cat in cats)
                {
                    AddAllowedCategoryName(cat, result);
                    try
                    {
                        var subs = cat.SubCategories;
                        if (subs != null)
                        {
                            foreach (Category subCat in subs) AddAllowedCategoryName(subCat, result);
                        }
                    }
                    catch { }
                }

                return result;
            }

            private static void AddAllowedCategoryName(Category cat, HashSet<string> allowedNames)
            {
                if (cat == null || allowedNames == null) return;
                string name;
                try { name = cat.Name ?? string.Empty; } catch { name = string.Empty; }
                if (string.IsNullOrWhiteSpace(name)) return;

                var trimmed = name.Trim();
                if (trimmed.StartsWith("<", StringComparison.OrdinalIgnoreCase)) return;
                if (trimmed.IndexOf("line style", StringComparison.OrdinalIgnoreCase) >= 0) return;

                bool canBind;
                try { canBind = cat.AllowsBoundParameters; } catch { canBind = false; }
                if (!canBind) return;

                allowedNames.Add(trimmed);
            }

            private static void AddDetailRow(DataTable dt, string rvtName, string rvtPath, string famName, string famCat,
                string pName, string paramKind, string isShared, string famGuid, string fileGuid, string res, string notes)
            {
                var r = dt.NewRow();
                r["RvtName"] = rvtName ?? string.Empty;
                r["RvtPath"] = rvtPath ?? string.Empty;
                r["FamilyName"] = famName ?? string.Empty;
                r["FamilyCategory"] = famCat ?? string.Empty;
                r["ParamName"] = pName ?? string.Empty;
                r["ParamKind"] = paramKind ?? string.Empty;
                r["IsShared"] = isShared ?? string.Empty;
                r["FamilyGuid"] = famGuid ?? string.Empty;
                r["FileGuid"] = fileGuid ?? string.Empty;
                r["Result"] = res ?? string.Empty;
                r["Notes"] = notes ?? string.Empty;
                dt.Rows.Add(r);
            }

            private static string GetFamilyParamKind(FamilyParameter fp)
            {
                if (fp == null) return "None";
                var isSharedFlag = false;
                try { isSharedFlag = fp.IsShared; } catch { }
                if (isSharedFlag) return "Shared";

                var idVal = 0;
                try { idVal = fp.Id.IntegerValue; } catch { idVal = 0; }
                if (idVal < 0) return "BuiltIn";
                return "Family";
            }

            private static string SafeParamElementName(Element pe)
            {
                try { return pe.Name; } catch { return string.Empty; }
            }

            private static bool TryGetFamilyParameterGuid(FamilyParameter fp, out Guid g)
            {
                g = Guid.Empty;
                if (fp == null) return false;

                var t = fp.GetType();
                var p = t.GetProperty("GUID", BindingFlags.Public | BindingFlags.Instance);
                if (p == null) return false;
                var v = p.GetValue(fp, null);
                if (v == null) return false;

                if (v is Guid gv)
                {
                    g = gv;
                    return g != Guid.Empty;
                }

                return false;
            }
        }


            private static string GetParamTypeName(Definition def)
            {
                if (def == null) return string.Empty;

                try
                {
                    var p = def.GetType().GetProperty("ParameterType", BindingFlags.Public | BindingFlags.Instance);
                    if (p != null)
                    {
                        var v = p.GetValue(def, null);
                        if (v != null) return v.ToString();
                    }
                }
                catch { }

                try
                {
                    var m = def.GetType().GetMethod("GetDataType", BindingFlags.Public | BindingFlags.Instance);
                    if (m != null)
                    {
                        var v = m.Invoke(def, null);
                        if (v != null) return v.ToString();
                    }
                }
                catch { }

                try
                {
                    var p2 = def.GetType().GetProperty("DataType", BindingFlags.Public | BindingFlags.Instance);
                    if (p2 != null)
                    {
                        var v = p2.GetValue(def, null);
                        if (v != null) return v.ToString();
                    }
                }
                catch { }

                return string.Empty;
            }


        public static DataTable CreateProjectTable() => Auditors.MakeFailureSummaryTable(1);
        public static DataTable CreateFamilyDetailTable() => Auditors.MakeFailureSummaryTable(2);

        public static DataTable CreateFamilyIndexTable()
        {
            var dt = new DataTable("FamilyIndex");
            dt.Columns.Add("RvtPath");
            dt.Columns.Add("RvtName");
            dt.Columns.Add("FamilyName");
            return dt;
        }
    }
}
