using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KKY_Tool_Revit.Exports;

namespace KKY_Tool_Revit.UI.Hub
{
    public sealed partial class UiBridgeExternalEvent
    {
        private readonly Stack<List<int>> _deleteOps = new Stack<List<int>>();
        private List<DupRowDto> _lastRows = new List<DupRowDto>();
        private static HashSet<int> _nestedSharedIds;

        private sealed class DupRowDto
        {
            public int ElementId { get; set; }
            public string Category { get; set; }
            public string Family { get; set; }
            public string Type { get; set; }
            public int ConnectedCount { get; set; }
            public string ConnectedIds { get; set; }
            public bool Candidate { get; set; }
            public bool Deleted { get; set; }
        }

        private void HandleDupRun(UIApplication app, object payload)
        {
            var uiDoc = app.ActiveUIDocument;
            if (uiDoc == null || uiDoc.Document == null)
            {
                SendToWeb("host:error", new { message = "활성 문서가 없습니다." });
                return;
            }

            var doc = uiDoc.Document;
            _nestedSharedIds = new HashSet<int>();

            try
            {
                var famCol = new FilteredElementCollector(doc);
                famCol.OfClass(typeof(FamilyInstance)).WhereElementIsNotElementType();
                foreach (Element o in famCol)
                {
                    var fi = o as FamilyInstance;
                    if (fi == null) continue;
                    try
                    {
                        var subs = fi.GetSubComponentIds();
                        if (subs == null) continue;
                        foreach (var sid in subs)
                        {
                            _nestedSharedIds.Add(sid.IntegerValue);
                        }
                    }
                    catch
                    {
                    }
                }
            }
            catch
            {
            }

            var tolFeet = 1.0 / 64.0;
            try
            {
                var tolObj = GetProp(payload, "tolFeet");
                if (tolObj != null) tolFeet = Math.Max(0.000001, Convert.ToDouble(tolObj));
            }
            catch
            {
            }

            var rows = new List<DupRowDto>();
            var total = 0;
            var groupsWithDup = 0;
            var candidates = 0;

            var collector = new FilteredElementCollector(doc);
            collector.WhereElementIsNotElementType();

            Func<double, long> q = x => (long)Math.Round(x / tolFeet);

            var buckets = new Dictionary<string, List<ElementId>>(StringComparer.Ordinal);
            var catCache = new Dictionary<int, string>();
            var famCache = new Dictionary<int, string>();
            var typCache = new Dictionary<int, string>();

            foreach (Element e in collector)
            {
                total += 1;
                if (ShouldSkipForQuantity(e)) continue;
                if (e == null || e.Category == null) continue;

                var center = TryGetCenter(e);
                if (center == null) continue;

                var catName = SafeCategoryName(e, catCache);
                var famName = SafeFamilyName(e, famCache);
                var typName = SafeTypeName(e, typCache);
                var lvl = TryGetLevelId(e);
                var oriKey = GetOrientationKey(e);

                var key = string.Concat(
                    catName, "|",
                    famName, "|",
                    typName, "|",
                    "O", oriKey, "|",
                    "L", lvl.ToString(), "|",
                    "Q(", q(center.X).ToString(), ",", q(center.Y).ToString(), ",", q(center.Z).ToString(), ")");

                if (!buckets.TryGetValue(key, out var list))
                {
                    list = new List<ElementId>();
                    buckets.Add(key, list);
                }

                list.Add(e.Id);
            }

            foreach (var kv in buckets)
            {
                var ids = kv.Value;
                if (ids.Count <= 1) continue;

                groupsWithDup += 1;

                foreach (var id in ids)
                {
                    var e = doc.GetElement(id);
                    if (e == null) continue;

                    var catName = SafeCategoryName(e, catCache);
                    var famName = SafeFamilyName(e, famCache);
                    var typName = SafeTypeName(e, typCache);

                    var connIds = ids
                        .Where(x => x.IntegerValue != id.IntegerValue)
                        .Select(x => x.IntegerValue.ToString())
                        .ToArray();

                    rows.Add(new DupRowDto
                    {
                        ElementId = id.IntegerValue,
                        Category = catName,
                        Family = famName,
                        Type = typName,
                        ConnectedCount = connIds.Length,
                        ConnectedIds = string.Join(", ", connIds),
                        Candidate = true,
                        Deleted = false
                    });

                    candidates += 1;
                }
            }

            _lastRows = rows;

            var wireRows = rows.Select(r => new
            {
                elementId = r.ElementId,
                category = r.Category,
                family = r.Family,
                type = r.Type,
                connectedCount = r.ConnectedCount,
                connectedIds = r.ConnectedIds,
                candidate = r.Candidate,
                deleted = r.Deleted
            }).ToList();

            SendToWeb("dup:list", wireRows);
            SendToWeb("dup:result", new { scan = total, groups = groupsWithDup, candidates });
        }

        private void HandleDuplicateSelect(UIApplication app, object payload)
        {
            var uiDoc = app.ActiveUIDocument;
            if (uiDoc == null) return;

            var idVal = SafeInt(GetProp(payload, "id"));
            if (idVal <= 0) return;

            var elId = new ElementId(idVal);
            var el = uiDoc.Document.GetElement(elId);
            if (el == null)
            {
                SendToWeb("host:warn", new { message = $"요소 {idVal} 을(를) 찾을 수 없습니다." });
                return;
            }

            try
            {
                uiDoc.Selection.SetElementIds(new List<ElementId> { elId });
            }
            catch
            {
            }

            var bb = GetBoundingBox(el);
            try
            {
                if (bb != null)
                {
                    var views = uiDoc.GetOpenUIViews();
                    var target = views.FirstOrDefault(v => v.ViewId.IntegerValue == uiDoc.ActiveView.Id.IntegerValue);
                    if (target != null) target.ZoomAndCenterRectangle(bb.Min, bb.Max);
                    else uiDoc.ShowElements(elId);
                }
                else
                {
                    uiDoc.ShowElements(elId);
                }
            }
            catch
            {
            }
        }

        private void HandleDuplicateDelete(UIApplication app, object payload)
        {
            var uiDoc = app.ActiveUIDocument;
            if (uiDoc == null || uiDoc.Document == null)
            {
                SendToWeb("revit:error", new { message = "활성 문서를 찾을 수 없습니다." });
                return;
            }

            var doc = uiDoc.Document;
            var ids = ExtractIds(payload);
            if (ids == null || ids.Count == 0)
            {
                SendToWeb("revit:error", new { message = "잘못된 요청입니다(id 누락/형식 오류)." });
                return;
            }

            var eidList = new List<ElementId>();
            foreach (var i in ids)
            {
                if (i > 0)
                {
                    var eid = new ElementId(i);
                    if (doc.GetElement(eid) != null) eidList.Add(eid);
                }
            }

            if (eidList.Count == 0)
            {
                SendToWeb("host:warn", new { message = "삭제할 유효한 요소가 없습니다." });
                return;
            }

            var actuallyDeleted = new List<int>();
            using (var t = new Transaction(doc, $"KKY Dup Delete ({eidList.Count})"))
            {
                t.Start();
                try
                {
                    doc.Delete(eidList);
                    t.Commit();
                }
                catch (Exception ex)
                {
                    t.RollBack();
                    SendToWeb("revit:error", new { message = $"삭제 실패({eidList.Count}개): {ex.Message}" });
                    return;
                }
            }

            foreach (var eid in eidList)
            {
                if (doc.GetElement(eid) == null)
                {
                    actuallyDeleted.Add(eid.IntegerValue);
                    var row = _lastRows.FirstOrDefault(r => r.ElementId == eid.IntegerValue);
                    if (row != null) row.Deleted = true;
                    SendToWeb("dup:deleted", new { id = eid.IntegerValue });
                }
            }

            if (actuallyDeleted.Count > 0)
            {
                _deleteOps.Push(actuallyDeleted);
            }
        }

        private void HandleDuplicateRestore(UIApplication app, object payload)
        {
            var uiDoc = app.ActiveUIDocument;
            if (uiDoc == null || uiDoc.Document == null)
            {
                SendToWeb("revit:error", new { message = "활성 문서를 찾을 수 없습니다." });
                return;
            }

            if (_deleteOps.Count == 0)
            {
                SendToWeb("host:warn", new { message = "되돌릴 수 있는 최신 삭제가 없습니다." });
                return;
            }

            var requestIds = ExtractIds(payload);
            var lastPack = _deleteOps.Peek();

            var same = requestIds != null &&
                       requestIds.Count == lastPack.Count &&
                       !requestIds.Except(lastPack).Any();

            if (!same)
            {
                SendToWeb("host:warn", new { message = "되돌리기는 직전 삭제 묶음만 가능합니다." });
                return;
            }

            try
            {
                var cmdId = RevitCommandId.LookupPostableCommandId(PostableCommand.Undo);
                if (cmdId == null)
                {
                    throw new InvalidOperationException("Undo 명령을 찾을 수 없습니다.");
                }

                uiDoc.Application.PostCommand(cmdId);
            }
            catch (Exception ex)
            {
                SendToWeb("revit:error", new { message = $"되돌리기 실패: {ex.Message}" });
                return;
            }

            _deleteOps.Pop();
            foreach (var i in lastPack)
            {
                var r = _lastRows.FirstOrDefault(x => x.ElementId == i);
                if (r != null) r.Deleted = false;
                SendToWeb("dup:restored", new { id = i });
            }
        }

        private void HandleDuplicateExport(UIApplication app, object payload = null)
        {
            if (_lastRows == null)
            {
                SendToWeb("host:warn", new { message = "내보낼 데이터가 없습니다." });
                return;
            }

            var token = GetProp(payload, "token") as string;

            try
            {
                var groupsCount = CountGroups(_lastRows);
                var desktop = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
                var todayToken = DateTime.Now.ToString("yyMMdd");
                var defaultFileName = $"{todayToken}_중복객체 검토결과_{groupsCount}개.xlsx";
                var defaultPath = Path.Combine(desktop, defaultFileName);

                var sfd = new Microsoft.Win32.SaveFileDialog
                {
                    Filter = "Excel Workbook (*.xlsx)|*.xlsx",
                    FileName = Path.GetFileName(defaultPath),
                    AddExtension = true,
                    DefaultExt = "xlsx",
                    OverwritePrompt = true
                };

                if (sfd.ShowDialog() != true)
                {
                    return;
                }

                var outPath = sfd.FileName;
                var doAutoFit = ParseExcelMode(payload);
                ExcelProgressReporter.Reset("dup:progress");
                DuplicateExport.Save(outPath, _lastRows.Cast<object>(), doAutoFit, "dup:progress");

                SendToWeb("dup:exported", new { path = outPath, ok = true, token });
            }
            catch (IOException)
            {
                var msg = "해당 파일이 열려 있어 저장에 실패했습니다. 엑셀에서 파일을 닫은 뒤 다시 시도해 주세요.";
                SendToWeb("dup:exported", new { ok = false, message = msg, token });
            }
            catch (Exception ex)
            {
                SendToWeb("dup:exported", new { ok = false, message = $"엑셀 내보내기에 실패했습니다: {ex.Message}", token });
            }
        }

        private static bool ShouldSkipForQuantity(Element e)
        {
            if (e == null) return true;
            if (e is ImportInstance) return true;

            try
            {
                if (e.Category == null) return true;
                if (e.Category.CategoryType != CategoryType.Model) return true;

                var n = (e.Category.Name ?? "").ToLowerInvariant();

                if (n.Contains("view") || n.Contains("viewport")) return true;
                if (n.Contains("level") || n.Contains("grid")) return true;
                if (n.Contains("reference plane") || n.Contains("work plane")) return true;
                if (n.Contains("scope box") || n.Contains("matchline")) return true;
                if (n.Contains("section line") || n.Contains("callout")) return true;
                if (n.Contains("sheet") || n.Contains("시트")) return true;
                if (n.Contains("line") || n.Contains("선")) return true;
                if (n.Contains("sketch") || n.Contains("area boundary")) return true;
                if (n.Contains("filled region") || n.Contains("detail item")) return true;
                if (n.Contains("detail line") || n.Contains("symbol")) return true;
                if (n.Contains("text note") || n.Contains("dimension")) return true;
                if (n.Contains("room tag") || n.Contains("space tag") || n.Contains("area tag")) return true;
                if (n.Contains("center line") || n.Contains("centerline") || n.Contains("중심선")) return true;
                if (n.StartsWith("analytical")) return true;
            }
            catch
            {
                return true;
            }

            var fi = e as FamilyInstance;
            if (fi != null)
            {
                try
                {
                    if (fi.SuperComponent != null) return true;
                    if (_nestedSharedIds != null && _nestedSharedIds.Contains(fi.Id.IntegerValue)) return true;
                }
                catch
                {
                }
            }

            try
            {
                var opts = new Options
                {
                    ComputeReferences = false,
                    IncludeNonVisibleObjects = false
                };

                if (!HasPositiveSolid(e, opts)) return true;
            }
            catch
            {
                return true;
            }

            return false;
        }

        private static bool HasPositiveSolid(Element el, Options opts)
        {
            var geom = el.Geometry(opts);
            if (geom == null) return false;
            foreach (var g in geom)
            {
                var s = g as Solid;
                if (s != null && s.Volume > 0) return true;

                var inst = g as GeometryInstance;
                if (inst != null)
                {
                    var instGeom = inst.GetInstanceGeometry();
                    if (instGeom != null)
                    {
                        foreach (var gi in instGeom)
                        {
                            var si = gi as Solid;
                            if (si != null && si.Volume > 0) return true;
                        }
                    }
                }
            }

            return false;
        }

        private static long QOri(double x)
        {
            return (long)Math.Round(x * 1000.0);
        }

        private static string GetOrientationKey(Element e)
        {
            try
            {
                var fi = e as FamilyInstance;
                if (fi != null)
                {
                    var mirrored = false;
                    var hand = false;
                    var facing = false;

                    try { mirrored = fi.Mirrored; } catch { }
                    try { hand = fi.HandFlipped; } catch { }
                    try { facing = fi.FacingFlipped; } catch { }

                    Transform t = null;
                    try { t = fi.GetTransform(); } catch { }

                    var keyParts = new List<string>
                    {
                        "M" + (mirrored ? "1" : "0"),
                        "H" + (hand ? "1" : "0"),
                        "F" + (facing ? "1" : "0")
                    };

                    if (t != null)
                    {
                        var ox = t.BasisX;
                        var oy = t.BasisY;
                        var oz = t.BasisZ;
                        keyParts.Add($"OX({QOri(ox.X)},{QOri(ox.Y)},{QOri(ox.Z)})");
                        keyParts.Add($"OY({QOri(oy.X)},{QOri(oy.Y)},{QOri(oy.Z)})");
                        keyParts.Add($"OZ({QOri(oz.X)},{QOri(oz.Y)},{QOri(oz.Z)})");
                    }

                    return string.Join("|", keyParts);
                }

                Location loc = null;
                try { loc = e.Location; } catch { }

                var lc = loc as LocationCurve;
                if (lc != null && lc.Curve != null)
                {
                    var c = lc.Curve;
                    XYZ dir = null;
                    try { dir = c.GetEndPoint(1) - c.GetEndPoint(0); } catch { }

                    if (dir != null)
                    {
                        var len = dir.GetLength();
                        if (len > 0.000001) dir = dir / len;
                        return $"LC({QOri(dir.X)},{QOri(dir.Y)},{QOri(dir.Z)})";
                    }
                }
            }
            catch
            {
            }

            return string.Empty;
        }

        private static string SafeCategoryName(Element e, Dictionary<int, string> cache)
        {
            if (e == null || e.Category == null) return "";
            var id = e.Category.Id.IntegerValue;
            if (cache.TryGetValue(id, out var s)) return s;
            s = e.Category.Name;
            cache[id] = s;
            return s;
        }

        private static string SafeFamilyName(Element e, Dictionary<int, string> cache)
        {
            var fi = e as FamilyInstance;
            if (fi == null || fi.Symbol == null || fi.Symbol.Family == null) return "";
            var id = fi.Symbol.Family.Id.IntegerValue;
            if (cache.TryGetValue(id, out var s)) return s;
            s = fi.Symbol.Family.Name;
            cache[id] = s;
            return s;
        }

        private static string SafeTypeName(Element e, Dictionary<int, string> cache)
        {
            var fi = e as FamilyInstance;
            if (fi != null && fi.Symbol != null)
            {
                var id = fi.Symbol.Id.IntegerValue;
                if (cache.TryGetValue(id, out var s)) return s;
                s = fi.Symbol.Name;
                cache[id] = s;
                return s;
            }

            return e.Name;
        }

        private static int TryGetLevelId(Element e)
        {
            try
            {
                var p = e.Parameter(BuiltInParameter.LEVEL_PARAM);
                if (p != null)
                {
                    var lvid = p.AsElementId();
                    if (lvid != null && lvid != ElementId.InvalidElementId) return lvid.IntegerValue;
                }
            }
            catch
            {
            }

            try
            {
                var pi = e.GetType().GetProperty("LevelId");
                if (pi != null)
                {
                    var id = pi.GetValue(e, null) as ElementId;
                    if (id != null && id != ElementId.InvalidElementId) return id.IntegerValue;
                }
            }
            catch
            {
            }

            return -1;
        }

        private static XYZ TryGetCenter(Element e)
        {
            if (e == null) return null;
            try
            {
                var loc = e.Location;
                if (loc is LocationPoint lp)
                {
                    return lp.Point;
                }

                if (loc is LocationCurve lc)
                {
                    var crv = lc.Curve;
                    if (crv != null) return crv.Evaluate(0.5, true);
                }
            }
            catch
            {
            }

            var bb = GetBoundingBox(e);
            if (bb != null)
            {
                return (bb.Min + bb.Max) * 0.5;
            }

            return null;
        }

        private static BoundingBoxXYZ GetBoundingBox(Element e)
        {
            try
            {
                var bb = e.get_BoundingBox(null);
                if (bb != null) return bb;
            }
            catch
            {
            }

            return null;
        }

        private static int SafeInt(object o)
        {
            if (o == null) return 0;
            try { return Convert.ToInt32(o); } catch { return 0; }
        }

        private static List<int> ExtractIds(object payload)
        {
            var result = new List<int>();

            var singleObj = GetProp(payload, "id");
            var v = SafeToInt(singleObj);
            if (v > 0)
            {
                result.Add(v);
                return result;
            }

            var arr = GetProp(payload, "ids");
            if (arr == null) return result;

            var enumerable = arr as IEnumerable;
            if (enumerable != null)
            {
                foreach (var o in enumerable)
                {
                    var iv = SafeToInt(o);
                    if (iv > 0) result.Add(iv);
                }
            }

            return result;
        }

        private static int SafeToInt(object o)
        {
            if (o == null) return 0;
            try
            {
                if (o is int i) return i;
                if (o is long l) return (int)l;
                if (o is double d) return (int)d;
                if (o is string s && int.TryParse(s, out var iv)) return iv;
            }
            catch
            {
            }

            return 0;
        }

        private int CountGroups(IEnumerable<DupRowDto> rows)
        {
            if (rows == null) return 0;

            var bucket = new HashSet<string>(StringComparer.Ordinal);
            foreach (var r in rows)
            {
                var id = r.ElementId.ToString();
                var cat = r.Category ?? "";
                var fam = r.Family ?? "";
                var typ = r.Type ?? "";
                var conStr = r.ConnectedIds ?? "";

                var cluster = new List<string>();
                if (!string.IsNullOrWhiteSpace(id)) cluster.Add(id);
                cluster.AddRange(SplitIds(conStr));

                var norm = cluster
                    .Where(x => !string.IsNullOrWhiteSpace(x))
                    .Select(x => x.Trim())
                    .Distinct()
                    .OrderBy(x => x)
                    .ToList();

                var clusterKey = norm.Count > 1 ? string.Join(",", norm) : "";
                var famOut = string.IsNullOrWhiteSpace(fam) ? (string.IsNullOrWhiteSpace(cat) ? "" : cat + " Type") : fam;
                var key = string.Join("|", new[] { cat, famOut, typ, clusterKey });
                bucket.Add(key);
            }

            return bucket.Count;
        }

        private IEnumerable<string> SplitIds(string s)
        {
            if (string.IsNullOrWhiteSpace(s)) return Array.Empty<string>();
            return s.Split(new[] { ',', ' ', ';', '|', '\t', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
        }
    }
}
