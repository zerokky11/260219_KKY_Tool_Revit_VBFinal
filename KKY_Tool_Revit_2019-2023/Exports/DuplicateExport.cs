using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using KKY_Tool_Revit.Infrastructure;

namespace KKY_Tool_Revit.Exports
{
    public class DupRowDto
    {
        public string Id { get; set; }
        public string Category { get; set; }
        public string Family { get; set; }
        public string Type { get; set; }
        public List<string> ConnectedIds { get; set; }
    }

    public static class DuplicateExport
    {
        public static string Save(IEnumerable rows, bool doAutoFit = false, string progressChannel = null)
        {
            var mapped = MapRows(rows);
            var dt = BuildSimpleTable(mapped);
            return ExcelCore.PickAndSaveXlsx("Duplicates (Simple)", dt, "Duplicates.xlsx", doAutoFit, progressChannel);
        }

        public static void Save(string outPath, IEnumerable rows, bool doAutoFit = false, string progressChannel = null)
        {
            Export(outPath, rows, doAutoFit, progressChannel);
        }

        public static void Export(string outPath, IEnumerable rows, bool doAutoFit = false, string progressChannel = null)
        {
            var mapped = MapRows(rows);
            var dt = BuildSimpleTable(mapped);
            ExcelCore.SaveStyledSimple(outPath, "Duplicates (Simple)", dt, "Group", doAutoFit, progressChannel);
        }

        private static List<DupRowDto> MapRows(IEnumerable rows)
        {
            var list = new List<DupRowDto>();
            if (rows == null) return list;

            foreach (var o in rows)
            {
                var it = new DupRowDto
                {
                    Id = ReadProp(o, "Id", "ID", "ElementId", "ElementID", "elementId"),
                    Category = ReadProp(o, "Category", "category"),
                    Family = ReadProp(o, "Family", "family"),
                    Type = ReadProp(o, "Type", "type"),
                    ConnectedIds = ReadList(o, "ConnectedIds", "connectedIds", "Links", "links", "connected", "Connected", "ConnectedElements")
                };
                list.Add(it);
            }

            return list;
        }

        private static DataTable BuildSimpleTable(List<DupRowDto> rows)
        {
            var dt = new DataTable("simple");
            dt.Columns.Add("Group");
            dt.Columns.Add("ID");
            dt.Columns.Add("Category");
            dt.Columns.Add("Family");
            dt.Columns.Add("Type");

            var groupList = GroupByLogic(rows);
            for (var i = 0; i < groupList.Count; i++)
            {
                var gName = $"Group{i + 1}";
                foreach (var r in groupList[i])
                {
                    var famOut = string.IsNullOrWhiteSpace(r.Family)
                        ? (string.IsNullOrWhiteSpace(r.Category) ? string.Empty : r.Category + " Type")
                        : r.Family;

                    var dr = dt.NewRow();
                    dr["Group"] = gName;
                    dr["ID"] = Nz(r.Id);
                    dr["Category"] = Nz(r.Category);
                    dr["Family"] = Nz(famOut);
                    dr["Type"] = Nz(r.Type);
                    dt.Rows.Add(dr);
                }
            }

            if (dt.Rows.Count == 0)
            {
                var dr = dt.NewRow();
                dr[0] = "오류가 없습니다.";
                dt.Rows.Add(dr);
            }

            return dt;
        }

        private static List<List<DupRowDto>> GroupByLogic(List<DupRowDto> items)
        {
            var buckets = new Dictionary<string, List<DupRowDto>>();

            foreach (var r in items)
            {
                var fam = string.IsNullOrWhiteSpace(r.Family)
                    ? (string.IsNullOrWhiteSpace(r.Category) ? string.Empty : r.Category + " Type")
                    : r.Family;
                var typ = string.IsNullOrWhiteSpace(r.Type) ? string.Empty : r.Type;
                var cat = string.IsNullOrWhiteSpace(r.Category) ? string.Empty : r.Category;

                var clusterSrc = new List<string>();
                if (!string.IsNullOrWhiteSpace(r.Id)) clusterSrc.Add(r.Id);
                if (r.ConnectedIds != null) clusterSrc.AddRange(r.ConnectedIds);

                var cluster = clusterSrc
                    .SelectMany(SplitIds)
                    .Where(x => !string.IsNullOrWhiteSpace(x))
                    .Select(x => x.Trim())
                    .Distinct()
                    .OrderBy(PadNum)
                    .ToList();

                var clusterKey = cluster.Count > 1 ? string.Join(",", cluster) : string.Empty;
                var key = string.Join("|", new[] { cat, fam, typ, clusterKey });
                if (!buckets.ContainsKey(key)) buckets[key] = new List<DupRowDto>();
                buckets[key].Add(r);
            }

            return buckets.Values.ToList();
        }

        private static IEnumerable<string> SplitIds(string s)
        {
            if (string.IsNullOrWhiteSpace(s)) return Array.Empty<string>();
            return s.Split(new[] { ',', ' ', ';', '|', '\t', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
        }

        private static string PadNum(string s)
        {
            if (int.TryParse(s, out var n)) return n.ToString("D10");
            return s;
        }

        private static string Nz(string s)
        {
            return string.IsNullOrWhiteSpace(s) ? string.Empty : s;
        }

        private static string ReadProp(object obj, params string[] names)
        {
            if (obj == null) return string.Empty;
            foreach (var nm in names)
            {
                if (string.IsNullOrEmpty(nm)) continue;
                var p = obj.GetType().GetProperty(nm);
                if (p == null) continue;
                var v = p.GetValue(obj, null);
                if (v != null) return v.ToString();
            }

            return string.Empty;
        }

        private static List<string> ReadList(object obj, params string[] names)
        {
            var res = new List<string>();
            if (obj == null) return res;

            foreach (var nm in names)
            {
                var p = obj.GetType().GetProperty(nm);
                if (p == null) continue;
                var v = p.GetValue(obj, null);
                if (v == null) continue;

                if (v is string sv)
                {
                    res.AddRange(SplitIds(sv));
                    break;
                }

                if (v is IEnumerable e && !(v is string))
                {
                    foreach (var x in e)
                    {
                        if (x != null) res.Add(x.ToString());
                    }
                    break;
                }
            }

            return res;
        }
    }
}
