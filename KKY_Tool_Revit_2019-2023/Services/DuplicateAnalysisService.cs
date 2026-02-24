using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Electrical;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.UI;

namespace KKY_Tool_Revit.Services
{
    public class DuplicateAnalysisService
    {
        public static List<Dictionary<string, object>> Run(UIApplication app)
        {
            var uidoc = app?.ActiveUIDocument;
            if (uidoc == null || uidoc.Document == null)
            {
                return new List<Dictionary<string, object>>();
            }

            var doc = uidoc.Document;
            var elems = new List<Element>();

            elems.AddRange(new FilteredElementCollector(doc)
                .OfClass(typeof(MEPCurve))
                .WhereElementIsNotElementType()
                .ToElements());

            elems.AddRange(new FilteredElementCollector(doc)
                .OfClass(typeof(FamilyInstance))
                .WhereElementIsNotElementType()
                .ToElements()
                .Where(HasAnyConnector));

            var groups = new Dictionary<string, List<Element>>(StringComparer.Ordinal);
            foreach (var e in elems)
            {
                var key = BuildGroupKey(e);
                if (!groups.ContainsKey(key)) groups[key] = new List<Element>();
                groups[key].Add(e);
            }

            var rows = new List<Dictionary<string, object>>();
            var gno = 1;

            foreach (var kv in groups)
            {
                var list = kv.Value;
                if (list == null || list.Count == 0) continue;

                var isCandidate = list.Count >= 2;
                foreach (var e in list)
                {
                    var cat = e.Category == null ? string.Empty : e.Category.Name;
                    var fam = string.Empty;
                    var typ = string.Empty;

                    if (e is FamilyInstance fi)
                    {
                        typ = fi.Symbol != null ? fi.Symbol.Name : string.Empty;
                        fam = fi.Symbol != null && fi.Symbol.Family != null ? fi.Symbol.Family.Name : string.Empty;
                    }
                    else
                    {
                        typ = e.Name;
                        var es = doc.GetElement(e.GetTypeId()) as ElementType;
                        if (es != null) fam = es.FamilyName;
                    }

                    var connected = GetConnectedOwnerIds(e);
                    var connectedIdsStr = string.Join(",", connected.Select(x => x.ToString()).ToArray());

                    var d = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase)
                    {
                        {"groupId", gno},
                        {"id", e.Id.IntegerValue.ToString()},
                        {"category", cat},
                        {"family", fam},
                        {"type", typ},
                        {"connectedCount", connected.Count},
                        {"connectedIds", connectedIdsStr},
                        {"candidate", isCandidate}
                    };
                    rows.Add(d);
                }

                gno += 1;
            }

            rows = rows.OrderBy(r => Convert.ToInt32(r["groupId"]))
                       .ThenBy(r => Convert.ToInt32(r["id"]))
                       .ToList();

            return rows;
        }

        public static bool Export(UIApplication app)
        {
            return false;
        }

        public static bool Restore(UIApplication app, Dictionary<string, object> payload)
        {
            return false;
        }

        private static bool HasAnyConnector(Element e)
        {
            try
            {
                if (e is MEPCurve mc)
                {
                    var cm = mc.ConnectorManager;
                    return cm != null && cm.Connectors != null && cm.Connectors.Size > 0;
                }

                var fi = e as FamilyInstance;
                if (fi != null && fi.MEPModel != null && fi.MEPModel.ConnectorManager != null)
                {
                    var cms = fi.MEPModel.ConnectorManager.Connectors;
                    return cms != null && cms.Size > 0;
                }
            }
            catch
            {
                // ignore
            }

            return false;
        }

        private static string BuildGroupKey(Element e)
        {
            var cat = e.Category == null ? string.Empty : e.Category.Name;
            var fam = string.Empty;
            var typ = string.Empty;
            var doc = e.Document;

            if (e is FamilyInstance fi)
            {
                typ = fi.Symbol != null ? fi.Symbol.Name : string.Empty;
                fam = fi.Symbol != null && fi.Symbol.Family != null ? fi.Symbol.Family.Name : string.Empty;
                var lp = e.Location as LocationPoint;
                if (lp == null || lp.Point == null)
                {
                    return $"{cat}|{fam}|{typ}|NOLOC";
                }

                var p = lp.Point;
                var keyp = $"{R2(p.X)}_{R2(p.Y)}_{R2(p.Z)}";
                return $"{cat}|{fam}|{typ}|{keyp}";
            }

            typ = e.Name;
            var es2 = doc.GetElement(e.GetTypeId()) as ElementType;
            if (es2 != null) fam = es2.FamilyName;

            var lc = e.Location as LocationCurve;
            if (lc == null || lc.Curve == null)
            {
                return $"{cat}|{fam}|{typ}|NOCURVE";
            }

            var c = lc.Curve;
            var s = c.GetEndPoint(0);
            var t = c.GetEndPoint(1);
            var k1 = $"{R2(s.X)}_{R2(s.Y)}_{R2(s.Z)}";
            var k2 = $"{R2(t.X)}_{R2(t.Y)}_{R2(t.Z)}";
            var a = new[] { k1, k2 };
            Array.Sort(a, StringComparer.Ordinal);
            return $"{cat}|{fam}|{typ}|{a[0]}|{a[1]}";
        }

        private static string R2(double v)
        {
            return Math.Round(v, 4, MidpointRounding.AwayFromZero).ToString("0.####");
        }

        private static HashSet<int> GetConnectedOwnerIds(Element e)
        {
            var setIds = new HashSet<int>();

            try
            {
                ConnectorManager cm = null;

                if (e is MEPCurve mc)
                {
                    cm = mc.ConnectorManager;
                }
                else
                {
                    var fi = e as FamilyInstance;
                    if (fi != null && fi.MEPModel != null)
                    {
                        cm = fi.MEPModel.ConnectorManager;
                    }
                }

                if (cm == null || cm.Connectors == null || cm.Connectors.Size == 0)
                {
                    return setIds;
                }

                foreach (var o in cm.Connectors)
                {
                    var c = o as Connector;
                    if (c == null) continue;
                    var refs = c.AllRefs;
                    if (refs == null || refs.Size == 0) continue;

                    foreach (var ro in refs)
                    {
                        var rc = ro as Connector;
                        if (rc == null || rc.Owner == null) continue;
                        var oid = rc.Owner.Id.IntegerValue;
                        if (oid != e.Id.IntegerValue) setIds.Add(oid);
                    }
                }
            }
            catch
            {
                // ignore
            }

            return setIds;
        }
    }
}
