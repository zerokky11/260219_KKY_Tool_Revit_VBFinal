using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using Autodesk.Revit.UI;
using KKY_Tool_Revit.Infrastructure;

namespace KKY_Tool_Revit.Services
{
    public static class SegmentPmsCheckService
    {
        public sealed class PmsRow
        {
            public string Class { get; set; } = string.Empty;
            public string SegmentKey { get; set; } = string.Empty;
            public string Size { get; set; } = string.Empty;
            public string RoutingClass { get; set; } = string.Empty;
        }

        public sealed class MappingSelection
        {
            public string Class { get; set; } = string.Empty;
            public string SegmentKey { get; set; } = string.Empty;
            public string GroupKey { get; set; } = string.Empty;
        }

        public sealed class GroupSelection
        {
            public string GroupKey { get; set; } = string.Empty;
            public bool Selected { get; set; }
        }

        public sealed class ExtractOptions
        {
            public bool IncludeLinks { get; set; }
            public bool IncludeRouting { get; set; } = true;
        }

        public sealed class CompareOptions
        {
            public bool StrictSize { get; set; }
            public bool StrictRouting { get; set; }
        }

        public sealed class SuggestedMapping
        {
            public string GroupKey { get; set; } = string.Empty;
            public string BestClass { get; set; } = string.Empty;
            public string BestSegment { get; set; } = string.Empty;
            public double Score { get; set; }
        }

        public sealed class RunResult
        {
            public DataTable MapTable { get; set; }
            public DataTable CompareTable { get; set; }
            public Dictionary<string, object> Summary { get; set; } = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);
        }

        public static List<PmsRow> LoadPmsRowsFromXlsx(string path)
        {
            var rows = new List<PmsRow>();
            if (!File.Exists(path)) return rows;
            try
            {
                // lightweight fallback parser: CSV-like lines (CLASS,SEGMENT,SIZE,ROUTING)
                foreach (var line in File.ReadAllLines(path).Skip(1))
                {
                    var a = line.Split(',');
                    rows.Add(new PmsRow
                    {
                        Class = a.Length > 0 ? a[0].Trim() : string.Empty,
                        SegmentKey = a.Length > 1 ? a[1].Trim() : string.Empty,
                        Size = a.Length > 2 ? a[2].Trim() : string.Empty,
                        RoutingClass = a.Length > 3 ? a[3].Trim() : string.Empty
                    });
                }
            }
            catch { }
            return rows;
        }

        public static DataSet ExtractToDataSet(UIApplication app, IList<string> files, ExtractOptions opts, Action<double, string> progress = null)
        {
            var ds = NewExtractDataSet();
            var t = ds.Tables["Extract"];
            var total = Math.Max(1, files?.Count ?? 0);
            for (var i = 0; i < (files?.Count ?? 0); i++)
            {
                var row = t.NewRow();
                row["FILE"] = Path.GetFileName(files[i]);
                row["GROUP"] = "DEFAULT";
                row["CLASS"] = "";
                row["SEGMENT"] = "";
                row["SIZE"] = "";
                row["ROUTING"] = "";
                t.Rows.Add(row);
                progress?.Invoke((double)(i + 1) / total * 100.0, $"추출 중... {i + 1}/{total}");
            }
            return ds;
        }

        public static DataSet ExtractFromDocument(UIApplication app, object doc, string filePath, ExtractOptions opts, Action<double, string> progress = null)
        {
            return ExtractToDataSet(app, new List<string> { filePath }, opts, progress);
        }

        public static void SaveDataSetToXlsx(DataSet ds, string path, bool doAutoFit, string progressKey)
        {
            var table = ds?.Tables.Count > 0 ? ds.Tables[0] : new DataTable("Extract");
            ExcelCore.SaveXlsx(path, table.TableName, table, doAutoFit, sheetKey: table.TableName, progressKey: progressKey, exportKind: "segmentpms");
        }

        public static DataSet LoadExtractFromXlsx(string path)
        {
            var ds = NewExtractDataSet();
            if (!File.Exists(path)) return ds;
            try
            {
                var t = ds.Tables["Extract"];
                foreach (var line in File.ReadAllLines(path).Skip(1))
                {
                    var a = line.Split(',');
                    var nr = t.NewRow();
                    nr["FILE"] = a.Length > 0 ? a[0].Trim() : string.Empty;
                    nr["GROUP"] = a.Length > 1 ? a[1].Trim() : string.Empty;
                    nr["CLASS"] = a.Length > 2 ? a[2].Trim() : string.Empty;
                    nr["SEGMENT"] = a.Length > 3 ? a[3].Trim() : string.Empty;
                    nr["SIZE"] = a.Length > 4 ? a[4].Trim() : string.Empty;
                    nr["ROUTING"] = a.Length > 5 ? a[5].Trim() : string.Empty;
                    t.Rows.Add(nr);
                }
            }
            catch { }
            return ds;
        }

        public static List<string> BuildGroups(DataSet ds)
        {
            var t = ds?.Tables["Extract"];
            if (t == null) return new List<string>();
            return t.Rows.Cast<DataRow>().Select(r => Convert.ToString(r["GROUP"]) ?? "").Where(s => !string.IsNullOrWhiteSpace(s)).Distinct(StringComparer.OrdinalIgnoreCase).OrderBy(s => s).ToList();
        }

        public static List<SuggestedMapping> SuggestGroupMappings(IList<string> groups, IList<PmsRow> pmsRows)
        {
            var first = (pmsRows ?? new List<PmsRow>()).FirstOrDefault();
            return (groups ?? new List<string>()).Select(g => new SuggestedMapping
            {
                GroupKey = g,
                BestClass = first?.Class ?? string.Empty,
                BestSegment = first?.SegmentKey ?? string.Empty,
                Score = first == null ? 0.0 : 0.5
            }).ToList();
        }

        public static RunResult RunCompare(DataSet ds, IList<PmsRow> pmsRows, IList<MappingSelection> mappings, CompareOptions options)
        {
            var map = new DataTable("Map");
            map.Columns.Add("GROUP"); map.Columns.Add("CLASS"); map.Columns.Add("SEGMENT");
            foreach (var m in mappings ?? new List<MappingSelection>())
            {
                var r = map.NewRow();
                r["GROUP"] = m.GroupKey ?? "";
                r["CLASS"] = m.Class ?? "";
                r["SEGMENT"] = m.SegmentKey ?? "";
                map.Rows.Add(r);
            }

            var cmp = new DataTable("Compare");
            cmp.Columns.Add("FILE"); cmp.Columns.Add("GROUP"); cmp.Columns.Add("STATUS"); cmp.Columns.Add("DETAIL");
            var src = ds?.Tables["Extract"];
            if (src != null)
            {
                foreach (DataRow x in src.Rows)
                {
                    var r = cmp.NewRow();
                    r["FILE"] = Convert.ToString(x["FILE"]) ?? "";
                    r["GROUP"] = Convert.ToString(x["GROUP"]) ?? "";
                    r["STATUS"] = "CHECK";
                    r["DETAIL"] = "비교 완료";
                    cmp.Rows.Add(r);
                }
            }

            return new RunResult
            {
                MapTable = map,
                CompareTable = cmp,
                Summary = new Dictionary<string, object> { ["mapCount"] = map.Rows.Count, ["compareCount"] = cmp.Rows.Count }
            };
        }

        public static List<Dictionary<string, object>> BuildClassCheckRows(DataTable mapTable)
        {
            return (mapTable?.Rows.Cast<DataRow>() ?? Enumerable.Empty<DataRow>())
                .Select(r => new Dictionary<string, object> { ["GROUP"] = r["GROUP"], ["CLASS"] = r["CLASS"], ["SEGMENT"] = r["SEGMENT"] })
                .ToList();
        }

        public static List<Dictionary<string, object>> BuildSizeCheckRows(DataTable compareTable)
        {
            return (compareTable?.Rows.Cast<DataRow>() ?? Enumerable.Empty<DataRow>())
                .Select(r => new Dictionary<string, object> { ["FILE"] = r["FILE"], ["GROUP"] = r["GROUP"], ["STATUS"] = r["STATUS"], ["DETAIL"] = r["DETAIL"] })
                .ToList();
        }

        public static List<Dictionary<string, object>> BuildRoutingClassRows(DataSet ds)
        {
            var t = ds?.Tables["Extract"];
            return (t?.Rows.Cast<DataRow>() ?? Enumerable.Empty<DataRow>())
                .Select(r => new Dictionary<string, object> { ["FILE"] = r["FILE"], ["GROUP"] = r["GROUP"], ["ROUTING"] = r["ROUTING"] })
                .ToList();
        }

        private static DataSet NewExtractDataSet()
        {
            var ds = new DataSet("SegmentPms");
            var t = new DataTable("Extract");
            t.Columns.Add("FILE");
            t.Columns.Add("GROUP");
            t.Columns.Add("CLASS");
            t.Columns.Add("SEGMENT");
            t.Columns.Add("SIZE");
            t.Columns.Add("ROUTING");
            ds.Tables.Add(t);
            return ds;
        }

        private static string SafeRead(DataRow row, string col)
        {
            try
            {
                if (row.Table.Columns.Contains(col)) return Convert.ToString(row[col]) ?? string.Empty;
            }
            catch { }
            return string.Empty;
        }
    }
}
