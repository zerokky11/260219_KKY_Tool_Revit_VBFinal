using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Electrical;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.UI;

namespace KKY_Tool_Revit.Services
{
    public class ConnectorDiagnosticsService
    {
        private sealed class ParamInfo
        {
            public bool Exists { get; set; }
            public bool HasValue { get; set; }
            public string Text { get; set; }
            public string CompareKey { get; set; }
        }

        public static List<string> LastDebug { get; set; }

        private static void Log(string msg)
        {
            if (LastDebug == null) LastDebug = new List<string>();
            LastDebug.Add($"{DateTime.Now:HH:mm:ss.fff} {msg}");
        }

        private sealed class TargetFilter
        {
            public Func<Element, bool> Evaluator { get; set; }
            public string PrimaryParam { get; set; } = string.Empty;
        }

        private static List<Dictionary<string, object>> Run(
            UIApplication app,
            double tolFt,
            string param,
            IEnumerable<string> extraParams,
            string targetFilter,
            bool excludeEndDummy,
            bool includeOkRows,
            Action<double, string> progress = null)
        {
            LastDebug = new List<string>();
            var uidoc = app?.ActiveUIDocument;
            if (uidoc == null || uidoc.Document == null)
            {
                Log("ActiveUIDocument 없음");
                return new List<Dictionary<string, object>>();
            }

            var doc = uidoc.Document;
            return RunCore(doc, tolFt, param, extraParams, targetFilter, excludeEndDummy, includeOkRows, progress, BuildFileLabel(doc));
        }

        private static List<Dictionary<string, object>> RunOnDocument(
            Document doc,
            double tolFt,
            string param,
            IEnumerable<string> extraParams,
            string targetFilter,
            bool excludeEndDummy,
            bool includeOkRows,
            Action<double, string> progress = null)
        {
            LastDebug = new List<string>();
            if (doc == null)
            {
                Log("Document 없음");
                return new List<Dictionary<string, object>>();
            }

            return RunCore(doc, tolFt, param, extraParams, targetFilter, excludeEndDummy, includeOkRows, progress, BuildFileLabel(doc));
        }

        public static List<Dictionary<string, object>> RunOnDocument(
            Document doc,
            double tol,
            string unit,
            string paramName,
            IEnumerable<string> extraParams,
            string targetFilter,
            bool excludeEndDummy,
            Action<double, string> progress = null)
        {
            var tolFt = ToTolFt(tol, unit);
            Debug.WriteLine($"[Connector] tol={tol}, unit={unit}, tolFt={tolFt}");
            return RunOnDocument(doc, tolFt, paramName, extraParams, targetFilter, excludeEndDummy, false, progress);
        }

        public static List<Dictionary<string, object>> Run(UIApplication app, double tol, string unit, string paramName, Action<double, string> progress = null)
        {
            Debug.WriteLine($"[Connector] tol={tol}, unit={unit}, tolFt={ToTolFt(tol, unit)}");
            return Run(app, ToTolFt(tol, unit), paramName, null, null, false, false, progress);
        }

        public static List<Dictionary<string, object>> Run(UIApplication app, double tol, string unit, string paramName, IEnumerable<string> extraParams, Action<double, string> progress = null)
        {
            Debug.WriteLine($"[Connector] tol={tol}, unit={unit}, tolFt={ToTolFt(tol, unit)}");
            return Run(app, ToTolFt(tol, unit), paramName, extraParams, null, false, false, progress);
        }

        public static List<Dictionary<string, object>> Run(UIApplication app, double tol, string unit, string paramName, IEnumerable<string> extraParams, string targetFilter, bool excludeEndDummy, Action<double, string> progress = null)
        {
            Debug.WriteLine($"[Connector] tol={tol}, unit={unit}, tolFt={ToTolFt(tol, unit)}");
            return Run(app, ToTolFt(tol, unit), paramName, extraParams, targetFilter, excludeEndDummy, false, progress);
        }

        private static List<Dictionary<string, object>> RunCore(
            Document doc,
            double tolFt,
            string param,
            IEnumerable<string> extraParams,
            string targetFilter,
            bool excludeEndDummy,
            bool includeOkRows,
            Action<double, string> progress,
            string fileLabel)
        {
            var rows = new List<Dictionary<string, object>>();
            if (doc == null) return rows;

            var normalizedExtras = NormalizeExtraParams(extraParams);
            var filter = ParseTargetFilter(targetFilter);

            Log($"DOC={fileLabel}, tolFt={tolFt:0.###}, param='{param}', extra={string.Join(",", normalizedExtras)}, targetFilter='{targetFilter}', excludeEndDummy={excludeEndDummy}, includeOkRows={includeOkRows}");

            var allElems = CollectElementsWithConnectors(doc, filter, excludeEndDummy);
            Log($"수집 요소(전체): {allElems.Count}");
            if (allElems.Count == 0)
            {
                Log("커넥터를 가진 요소가 없습니다.");
                return rows;
            }

            var elemConns = new Dictionary<int, List<Connector>>();
            foreach (var el in allElems)
            {
                elemConns[el.Id.IntegerValue] = GetConnectors(el);
            }

            var allConnPoints = new List<Tuple<int, XYZ, Connector>>();
            foreach (var kv in elemConns)
            {
                foreach (var c in kv.Value)
                {
                    try
                    {
                        var org = c.Origin;
                        if (org != null) allConnPoints.Add(Tuple.Create(kv.Key, org, c));
                    }
                    catch
                    {
                    }
                }
            }

            var totalElem = Math.Max(1, allElems.Count);
            for (var i = 0; i < allElems.Count; i++)
            {
                var el = allElems[i];
                var baseId = el.Id.IntegerValue;
                var conns = elemConns[baseId];
                var connTotal = Math.Max(1, conns.Count);

                for (var j = 0; j < conns.Count; j++)
                {
                    var c = conns[j];
                    if (c == null) continue;

                    try
                    {
                        Element found = null;
                        var distFt = double.NaN;
                        var connType = "NONE";

                        if (c.IsConnected)
                        {
                            foreach (Connector r in c.AllRefs)
                            {
                                if (r == null || r.Owner == null) continue;
                                if (r.Owner.Id.IntegerValue == baseId) continue;
                                if (r.Owner is MEPSystem) continue;
                                found = r.Owner;
                                try { distFt = c.Origin.DistanceTo(r.Origin); } catch { distFt = 0; }
                                connType = "CONNECTED";
                                break;
                            }
                        }

                        if (found == null)
                        {
                            var nearest = FindNearestExternal(baseId, c, allConnPoints, tolFt);
                            if (nearest != null)
                            {
                                found = doc.GetElement(new ElementId(nearest.Item1));
                                distFt = nearest.Item3;
                                connType = "PROXIMITY";
                            }
                        }

                        if (found == null) continue;

                        var p1 = ReadParamInfo(el, param);
                        var p2 = ReadParamInfo(found, param);

                        var paramCompare = CompareParamInfo(p1, p2);
                        var status = "OK";
                        if (connType == "PROXIMITY" && !double.IsNaN(distFt) && Math.Abs(distFt) < 1.0e-9) status = "ERROR";
                        if (paramCompare == "Mismatch") status = "ERROR";

                        if (includeOkRows || !string.Equals(status, "OK", StringComparison.OrdinalIgnoreCase))
                        {
                            var extras1 = ReadExtraParams(el, normalizedExtras);
                            var extras2 = ReadExtraParams(found, normalizedExtras);
                            rows.Add(BuildRow(el, found, distFt * 12.0, connType, param, p1.Text, p2.Text, status, paramCompare, normalizedExtras, extras1, extras2, fileLabel));
                        }

                        if (progress != null)
                        {
                            var baseFrac = (double)i / totalElem;
                            var withinFrac = ((double)(j + 1) / connTotal) / totalElem;
                            var pct = Math.Round((baseFrac + withinFrac) * 1000.0) / 10.0;
                            if ((i < totalElem - 1 || j < connTotal - 1) && pct >= 100.0) pct = 99.9;
                            progress(pct, $"커넥터 진단 중... ({i + 1}/{totalElem})  커넥터 {j + 1}/{connTotal}");
                        }
                    }
                    catch (Exception ex)
                    {
                        rows.Add(new Dictionary<string, object>(StringComparer.Ordinal)
                        {
                            ["File"] = fileLabel,
                            ["Id1"] = baseId.ToString(),
                            ["Id2"] = "",
                            ["Category1"] = SafeCategoryName(el),
                            ["Category2"] = "",
                            ["Family1"] = GetFamilyName(el),
                            ["Family2"] = "",
                            ["Distance (inch)"] = "",
                            ["ConnectionType"] = "ERROR",
                            ["ParamName"] = param,
                            ["Value1"] = "",
                            ["Value2"] = "",
                            ["ParamCompare"] = "N/A",
                            ["Status"] = "ERROR",
                            ["ErrorMessage"] = ex.Message
                        });
                    }
                }
            }

            progress?.Invoke(100.0, "완료");

            rows = rows.OrderBy(r => ToDouble(GetDictValue(r, "Distance (inch)")))
                       .ThenBy(r => ToInt(GetDictValue(r, "Id1")))
                       .ThenBy(r => ToInt(GetDictValue(r, "Id2")))
                       .ToList();
            return rows;
        }

        private static Tuple<int, XYZ, double> FindNearestExternal(int baseId, Connector src, List<Tuple<int, XYZ, Connector>> allConnPoints, double tolFt)
        {
            if (src == null) return null;
            XYZ o;
            try { o = src.Origin; } catch { return null; }
            if (o == null) return null;

            Tuple<int, XYZ, double> best = null;
            var bestDist = double.MaxValue;
            foreach (var x in allConnPoints)
            {
                if (x.Item1 == baseId) continue;
                var d = o.DistanceTo(x.Item2);
                if (d <= tolFt && d < bestDist)
                {
                    bestDist = d;
                    best = Tuple.Create(x.Item1, x.Item2, d);
                }
            }
            return best;
        }

        private static List<Element> CollectElementsWithConnectors(Document doc, TargetFilter filter, bool excludeEndDummy)
        {
            var elems = new List<Element>();

            foreach (FamilyInstance fi in new FilteredElementCollector(doc).OfClass(typeof(FamilyInstance)))
            {
                try
                {
                    if (fi?.MEPModel?.ConnectorManager?.Connectors != null && fi.MEPModel.ConnectorManager.Connectors.Cast<Connector>().Any())
                    {
                        if (IsElementAllowed(fi, filter, excludeEndDummy)) elems.Add(fi);
                    }
                }
                catch
                {
                }
            }

            var cats = new[]
            {
                BuiltInCategory.OST_PipeCurves, BuiltInCategory.OST_DuctCurves, BuiltInCategory.OST_CableTray, BuiltInCategory.OST_Conduit,
                BuiltInCategory.OST_PipeFitting, BuiltInCategory.OST_DuctFitting, BuiltInCategory.OST_CableTrayFitting, BuiltInCategory.OST_ConduitFitting,
                BuiltInCategory.OST_PipeAccessory, BuiltInCategory.OST_DuctAccessory
            };

            foreach (var cat in cats)
            {
                foreach (Element el in new FilteredElementCollector(doc).OfCategory(cat).WhereElementIsNotElementType())
                {
                    if (HasConnectors(el) && IsElementAllowed(el, filter, excludeEndDummy)) elems.Add(el);
                }
            }

            return elems.GroupBy(e => e.Id.IntegerValue).Select(g => g.First()).ToList();
        }

        private static bool IsElementAllowed(Element el, TargetFilter filter, bool excludeEndDummy)
        {
            if (el == null) return false;
            if (excludeEndDummy)
            {
                var fn = GetFamilyName(el);
                if (!string.IsNullOrWhiteSpace(fn) && fn.IndexOf("End Dummy", StringComparison.OrdinalIgnoreCase) >= 0) return false;
            }
            if (filter?.Evaluator != null) return filter.Evaluator(el);
            return true;
        }

        private static bool HasConnectors(Element el)
        {
            try
            {
                if (el is FamilyInstance fi && fi.MEPModel?.ConnectorManager?.Connectors != null)
                    return fi.MEPModel.ConnectorManager.Connectors.Cast<Connector>().Any();
                if (el is MEPCurve mc && mc.ConnectorManager?.Connectors != null)
                    return mc.ConnectorManager.Connectors.Cast<Connector>().Any();
            }
            catch
            {
            }
            return false;
        }

        private static List<Connector> GetConnectors(Element el)
        {
            try
            {
                if (el is FamilyInstance fi && fi.MEPModel?.ConnectorManager?.Connectors != null)
                    return fi.MEPModel.ConnectorManager.Connectors.Cast<Connector>().ToList();
                if (el is MEPCurve mc && mc.ConnectorManager?.Connectors != null)
                    return mc.ConnectorManager.Connectors.Cast<Connector>().ToList();
            }
            catch
            {
            }
            return new List<Connector>();
        }

        private static ParamInfo ReadParamInfo(Element el, string paramName)
        {
            var info = new ParamInfo { Exists = false, HasValue = false, Text = string.Empty, CompareKey = string.Empty };
            if (el == null || string.IsNullOrWhiteSpace(paramName)) return info;
            try
            {
                var p = el.LookupParameter(paramName);
                if (p == null) return info;
                info.Exists = true;
                var txt = p.AsString();
                if (string.IsNullOrWhiteSpace(txt)) txt = p.AsValueString();
                if (string.IsNullOrWhiteSpace(txt) && p.StorageType == StorageType.Double)
                {
                    txt = p.AsDouble().ToString("0.########", CultureInfo.InvariantCulture);
                }
                info.Text = txt ?? string.Empty;
                info.HasValue = !string.IsNullOrWhiteSpace(info.Text);
                info.CompareKey = (info.Text ?? string.Empty).Trim().ToLowerInvariant();
            }
            catch
            {
            }
            return info;
        }

        private static string CompareParamInfo(ParamInfo p1, ParamInfo p2)
        {
            if ((p1 == null || !p1.Exists) && (p2 == null || !p2.Exists)) return "N/A";
            if (p1 == null || p2 == null || !p1.Exists || !p2.Exists) return "Mismatch";
            if (!p1.HasValue && !p2.HasValue) return "N/A";
            return string.Equals(p1.CompareKey, p2.CompareKey, StringComparison.Ordinal) ? "OK" : "Mismatch";
        }

        private static Dictionary<string, string> ReadExtraParams(Element el, IList<string> names)
        {
            var d = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            if (el == null || names == null) return d;
            foreach (var n in names)
            {
                try
                {
                    var p = el.LookupParameter(n);
                    var t = p == null ? "" : (p.AsString() ?? p.AsValueString() ?? "");
                    d[n] = t;
                }
                catch
                {
                    d[n] = "";
                }
            }
            return d;
        }

        private static Dictionary<string, object> BuildRow(
            Element e1,
            Element e2,
            double distInch,
            string connType,
            string param,
            string v1,
            string v2,
            string status,
            string paramCompare,
            IList<string> extraNames,
            Dictionary<string, string> extraVals1,
            Dictionary<string, string> extraVals2,
            string fileLabel)
        {
            var row = new Dictionary<string, object>(StringComparer.Ordinal)
            {
                ["File"] = fileLabel,
                ["Id1"] = e1 != null ? e1.Id.IntegerValue.ToString() : "0",
                ["Id2"] = e2 != null ? "," + e2.Id.IntegerValue : "",
                ["Category1"] = SafeCategoryName(e1),
                ["Category2"] = SafeCategoryName(e2),
                ["Family1"] = GetFamilyName(e1),
                ["Family2"] = GetFamilyName(e2),
                ["Distance (inch)"] = double.IsNaN(distInch) ? (object)"" : distInch,
                ["ConnectionType"] = connType,
                ["ParamName"] = param,
                ["Value1"] = v1 ?? "",
                ["Value2"] = v2 ?? "",
                ["ParamCompare"] = paramCompare,
                ["Status"] = status,
                ["ErrorMessage"] = ""
            };

            if (extraNames != null)
            {
                foreach (var name in extraNames)
                {
                    row[$"{name}(ID1)"] = extraVals1 != null && extraVals1.ContainsKey(name) ? extraVals1[name] : "";
                    row[$"{name}(ID2)"] = extraVals2 != null && extraVals2.ContainsKey(name) ? extraVals2[name] : "";
                }
            }

            return row;
        }

        private static string SafeCategoryName(Element e)
        {
            try { return e?.Category?.Name ?? ""; } catch { return ""; }
        }

        private static string GetFamilyName(Element e)
        {
            if (e == null) return "";
            try
            {
                if (e is FamilyInstance fi && fi.Symbol?.Family != null) return fi.Symbol.Family.Name;
                var et = e.Document.GetElement(e.GetTypeId()) as ElementType;
                if (et != null) return et.FamilyName;
            }
            catch
            {
            }
            return "";
        }

        private static string BuildFileLabel(Document doc)
        {
            try
            {
                if (!string.IsNullOrWhiteSpace(doc.PathName)) return System.IO.Path.GetFileName(doc.PathName);
                return doc.Title ?? "(Current)";
            }
            catch
            {
                return "(Current)";
            }
        }

        private static List<string> NormalizeExtraParams(IEnumerable<string> extraParams)
        {
            var list = new List<string>();
            var dedup = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            if (extraParams == null) return list;
            foreach (var p in extraParams)
            {
                if (string.IsNullOrWhiteSpace(p)) continue;
                var t = p.Trim();
                if (dedup.Add(t)) list.Add(t);
            }
            return list;
        }

        private static TargetFilter ParseTargetFilter(string raw)
        {
            if (string.IsNullOrWhiteSpace(raw)) return null;
            var tokens = raw.Split(new[] { ',', ';' }, StringSplitOptions.RemoveEmptyEntries)
                            .Select(x => x.Trim())
                            .Where(x => x.Length > 0)
                            .ToList();
            if (tokens.Count == 0) return null;

            return new TargetFilter
            {
                Evaluator = e =>
                {
                    var cat = SafeCategoryName(e);
                    var fam = GetFamilyName(e);
                    var tname = e?.Name ?? "";
                    var bag = (cat + "|" + fam + "|" + tname).ToLowerInvariant();
                    return tokens.Any(t => bag.Contains(t.ToLowerInvariant()));
                },
                PrimaryParam = string.Empty
            };
        }

        private static object GetDictValue(Dictionary<string, object> d, string key)
        {
            if (d == null) return null;
            return d.TryGetValue(key, out var v) ? v : null;
        }

        private static double ToDouble(object o)
        {
            if (o == null) return double.MaxValue;
            try { return Convert.ToDouble(o, CultureInfo.InvariantCulture); }
            catch { return double.MaxValue; }
        }

        private static int ToInt(object o)
        {
            if (o == null) return 0;
            try
            {
                var s = Convert.ToString(o);
                if (string.IsNullOrWhiteSpace(s)) return 0;
                s = s.Replace(",", "").Trim();
                return int.TryParse(s, out var v) ? v : 0;
            }
            catch { return 0; }
        }

        public static double ToTolFt(double tol, string unit)
        {
            var u = (unit ?? "inch").Trim().ToLowerInvariant();
            if (u == "ft" || u == "feet") return tol;
            if (u == "mm") return tol / 304.8;
            if (u == "m") return tol / 0.3048;
            return tol / 12.0; // inch default
        }
    }
}
