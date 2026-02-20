using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using Autodesk.Revit.UI;
using KKY_Tool_Revit.Infrastructure;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace KKY_Tool_Revit.Services
{
    public static class SegmentPmsCheckService
    {
        public const string TableMeta = "Extract_Meta";
        public const string TableFiles = "Extract_Files";
        public const string TableRules = "Extract_Rules";
        public const string TableSizes = "Extract_Sizes";
        public const string TableRouting = "Extract_Routing";

        public sealed class PmsRow
        {
            public string Class { get; set; } = string.Empty;
            public string SegmentKey { get; set; } = string.Empty;
            public double NdMm { get; set; }
            public double IdMm { get; set; }
            public double OdMm { get; set; }
        }

        public sealed class LoadPmsResult
        {
            public DataTable Table { get; set; }
            public List<PmsRow> Rows { get; set; } = new List<PmsRow>();
            public List<string> Errors { get; set; } = new List<string>();
        }

        public sealed class MappingSelection
        {
            public string File { get; set; } = string.Empty;
            public string PipeTypeName { get; set; } = string.Empty;
            public int RuleIndex { get; set; }
            public int SegmentId { get; set; }
            public string SegmentKey { get; set; } = string.Empty;
            public string SelectedClass { get; set; } = string.Empty;
            public string SelectedPmsSegment { get; set; } = string.Empty;
            public string MappingSource { get; set; } = string.Empty;

            // UI backward compatibility
            public string Class { get => SelectedClass; set => SelectedClass = value ?? string.Empty; }
            public string GroupKey { get => SegmentKey; set => SegmentKey = value ?? string.Empty; }
        }

        public sealed class GroupSelection
        {
            public string GroupKey { get; set; } = string.Empty;
            public string SelectedClass { get; set; } = string.Empty;
            public string SelectedPmsSegment { get; set; } = string.Empty;
            public string SelectionSource { get; set; } = string.Empty;
        }

        public sealed class ExtractOptions
        {
            public int NdRound { get; set; } = 3;
            public bool DetachFromCentral { get; set; } = true;
            public bool OpenReadOnly { get; set; } = true;
            public double ToleranceMm { get; set; } = 0.01;
        }

        public sealed class CompareOptions
        {
            public int NdRound { get; set; } = 3;
            public double TolMm { get; set; } = 0.01;
            public bool ClassMatch { get; set; }

            // backward compatibility
            public bool StrictSize { get => TolMm <= 0.001; set { if (value) TolMm = 0.001; } }
            public bool StrictRouting { get => ClassMatch; set => ClassMatch = value; }
        }

        public sealed class SuggestedMapping
        {
            public string File { get; set; } = string.Empty;
            public string PipeTypeName { get; set; } = string.Empty;
            public int RuleIndex { get; set; }
            public int SegmentId { get; set; }
            public string SegmentKey { get; set; } = string.Empty;
            public string PmsClass { get; set; } = string.Empty;
            public string PmsSegmentKey { get; set; } = string.Empty;
            public double Score { get; set; }

            // UI backward compatibility
            public string GroupKey { get => SegmentKey; set => SegmentKey = value ?? string.Empty; }
            public string BestClass { get => PmsClass; set => PmsClass = value ?? string.Empty; }
            public string BestSegment { get => PmsSegmentKey; set => PmsSegmentKey = value ?? string.Empty; }
        }

        public sealed class MappingUsage
        {
            public string File { get; set; } = string.Empty;
            public string PipeTypeName { get; set; } = string.Empty;
            public int RuleIndex { get; set; }
            public int SegmentId { get; set; }
            public string SegmentKey { get; set; } = string.Empty;
        }

        public sealed class MappingGroup
        {
            public string GroupKey { get; set; } = string.Empty;
            public string DisplayKey { get; set; } = string.Empty;
            public string NormalizedKey { get; set; } = string.Empty;
            public List<MappingUsage> Usages { get; set; } = new List<MappingUsage>();
            public string SuggestedClass { get; set; } = string.Empty;
            public string SuggestedSegmentKey { get; set; } = string.Empty;
            public int FileCount { get; set; }
            public int PipeTypeCount { get; set; }
            public string UsageSummary { get; set; } = string.Empty;
        }

        public sealed class RunResult
        {
            public DataTable MapTable { get; set; }
            public DataTable RevitSizeTable { get; set; }
            public DataTable PmsSizeTable { get; set; }
            public DataTable CompareTable { get; set; }
            public DataTable ErrorTable { get; set; }
            public DataTable SummaryTable { get; set; }
            public Dictionary<string, object> Summary { get; set; } = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);
        }

        public static DataSet ExtractToDataSet(UIApplication app,
            IList<string> files,
            ExtractOptions options,
            Action<double, string> progress = null)
        {
            var ds = BuildExtractDataSet(options ?? new ExtractOptions());
            var fileTable = ds.Tables[TableFiles];
            var rules = ds.Tables[TableRules];

            var list = files ?? new List<string>();
            var total = Math.Max(1, list.Count);
            for (var i = 0; i < list.Count; i++)
            {
                var p = list[i] ?? string.Empty;
                var fr = fileTable.NewRow();
                fr["File"] = p;
                fr["FileName"] = SafeFileName(p);
                fr["ExtractedAt"] = DateTime.Now.ToString("s", CultureInfo.InvariantCulture);
                fileTable.Rows.Add(fr);

                var rr = rules.NewRow();
                rr["File"] = p;
                rr["FileName"] = SafeFileName(p);
                rr["PipeTypeName"] = string.Empty;
                rr["RuleIndex"] = 0;
                rr["RevitSegmentKey"] = string.Empty;
                rr["SegmentId"] = 0;
                rules.Rows.Add(rr);

                progress?.Invoke((i + 1) * 100.0 / total, $"추출 중... {i + 1}/{total}");
            }

            return ds;
        }

        public static DataSet ExtractFromDocument(UIApplication app,
            object doc,
            string filePath,
            ExtractOptions options,
            Action<double, string> progress = null)
        {
            return ExtractToDataSet(app, new List<string> { filePath }, options, progress);
        }

        public static void SaveDataSetToXlsx(DataSet ds, string path, bool doAutoFit, string progressKey)
        {
            if (ds == null || string.IsNullOrWhiteSpace(path)) return;
            var sheets = new List<KeyValuePair<string, DataTable>>();
            foreach (DataTable t in ds.Tables)
            {
                sheets.Add(new KeyValuePair<string, DataTable>(t.TableName, t));
            }

            if (sheets.Count == 0)
            {
                var dt = new DataTable("Empty");
                dt.Columns.Add("Message");
                var r = dt.NewRow();
                r[0] = "데이터가 없습니다.";
                dt.Rows.Add(r);
                sheets.Add(new KeyValuePair<string, DataTable>("Empty", dt));
            }

            ExcelCore.SaveXlsxMulti(path, sheets, doAutoFit, progressKey);
        }

        public static DataSet LoadExtractFromXlsx(string path)
        {
            // 환경 제한상 xlsx 역직렬화 유틸이 없어, 스키마 중심 복원
            var ds = BuildExtractDataSet(new ExtractOptions());
            var fallback = Path.ChangeExtension(path ?? string.Empty, ".csv");
            if (!File.Exists(fallback)) return ds;

            try
            {
                var rules = ds.Tables[TableRules];
                foreach (var line in File.ReadAllLines(fallback).Skip(1))
                {
                    var a = line.Split(',');
                    var rr = rules.NewRow();
                    rr["File"] = a.Length > 0 ? a[0].Trim() : string.Empty;
                    rr["FileName"] = SafeFileName(Convert.ToString(rr["File"]));
                    rr["PipeTypeName"] = a.Length > 1 ? a[1].Trim() : string.Empty;
                    rr["RuleIndex"] = a.Length > 2 && int.TryParse(a[2], out var ri) ? ri : 0;
                    rr["RevitSegmentKey"] = a.Length > 3 ? a[3].Trim() : string.Empty;
                    rr["SegmentId"] = a.Length > 4 && int.TryParse(a[4], out var sid) ? sid : 0;
                    rules.Rows.Add(rr);
                }
            }
            catch
            {
            }

            return ds;
        }

        public static List<PmsRow> LoadPmsRowsFromXlsx(string path)
        {
            return LoadPmsExcel(path).Rows;
        }

        public static LoadPmsResult LoadPmsExcel(string path)
        {
            var result = new LoadPmsResult
            {
                Table = BuildPmsTableSkeleton()
            };

            if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
            {
                result.Errors.Add("PMS 파일이 없습니다.");
                return result;
            }

            try
            {
                using (var fs = File.Open(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    IWorkbook wb = new XSSFWorkbook(fs);
                    var sh = wb.GetSheetAt(0);
                    if (sh == null)
                    {
                        result.Errors.Add("PMS 시트를 찾을 수 없습니다.");
                        return result;
                    }

                    var header = sh.GetRow(sh.FirstRowNum);
                    if (!ValidatePmsHeader(header))
                    {
                        result.Errors.Add("PMS 헤더가 올바르지 않습니다. (CLASS, SEGMENT, ND_mm, ID_mm, OD_mm 필요)");
                        return result;
                    }

                    for (var r = sh.FirstRowNum + 1; r <= sh.LastRowNum; r++)
                    {
                        var row = sh.GetRow(r);
                        if (row == null) continue;

                        var cls = CellStr(row.GetCell(0));
                        var seg = CellStr(row.GetCell(1));
                        var nd = CellDbl(row.GetCell(2));
                        var id = CellDbl(row.GetCell(3));
                        var od = CellDbl(row.GetCell(4));

                        if (string.IsNullOrWhiteSpace(cls) && string.IsNullOrWhiteSpace(seg)) continue;

                        var item = new PmsRow
                        {
                            Class = cls,
                            SegmentKey = NormalizeSegmentGroupKey(seg),
                            NdMm = nd,
                            IdMm = id,
                            OdMm = od
                        };
                        result.Rows.Add(item);

                        var tr = result.Table.NewRow();
                        tr["CLASS"] = item.Class;
                        tr["SEGMENT"] = item.SegmentKey;
                        tr["ND_mm"] = item.NdMm;
                        tr["ID_mm"] = item.IdMm;
                        tr["OD_mm"] = item.OdMm;
                        result.Table.Rows.Add(tr);
                    }
                }
            }
            catch (Exception ex)
            {
                result.Errors.Add(ex.Message);
            }

            return result;
        }

        public static void ExportPmsTemplateXlsx(string path)
        {
            var dt = BuildPmsTableSkeleton();
            ExcelCore.SaveXlsx(path, "PMS_TEMPLATE", dt, autoFit: false, sheetKey: "segmentpms", progressKey: "segmentpms:progress", exportKind: "segmentpms");
        }

        public static List<MappingGroup> BuildGroups(DataSet extractData)
        {
            var groups = new Dictionary<string, MappingGroup>(StringComparer.OrdinalIgnoreCase);
            var rules = extractData?.Tables[TableRules];
            if (rules == null) return new List<MappingGroup>();

            foreach (DataRow r in rules.Rows)
            {
                var key = NormalizeSegmentGroupKey(SafeStr(r, "RevitSegmentKey"));
                if (!groups.TryGetValue(key, out var g))
                {
                    g = new MappingGroup
                    {
                        GroupKey = key,
                        DisplayKey = key,
                        NormalizedKey = key
                    };
                    groups[key] = g;
                }

                g.Usages.Add(new MappingUsage
                {
                    File = SafeStr(r, "File"),
                    PipeTypeName = SafeStr(r, "PipeTypeName"),
                    RuleIndex = SafeIntObj(r["RuleIndex"]),
                    SegmentId = SafeIntObj(r["SegmentId"]),
                    SegmentKey = key
                });
            }

            foreach (var g in groups.Values)
            {
                g.FileCount = g.Usages.Select(u => u.File).Distinct(StringComparer.OrdinalIgnoreCase).Count();
                g.PipeTypeCount = g.Usages.Select(u => u.PipeTypeName).Distinct(StringComparer.OrdinalIgnoreCase).Count();
                g.UsageSummary = $"Files:{g.FileCount}, PipeTypes:{g.PipeTypeCount}, Rules:{g.Usages.Count}";
            }

            return groups.Values.OrderBy(x => x.GroupKey, StringComparer.OrdinalIgnoreCase).ToList();
        }

        // UI backward compatibility overload
        public static List<string> BuildGroups(DataSet ds, bool flat)
        {
            return BuildGroups(ds).Select(g => g.GroupKey).ToList();
        }

        public static List<SuggestedMapping> SuggestMappings(DataSet ds, List<PmsRow> pmsRows)
        {
            var groups = BuildGroups(ds);
            return SuggestGroupMappings(groups, pmsRows);
        }

        public static List<SuggestedMapping> SuggestGroupMappings(IList<string> groups, IList<PmsRow> pmsRows)
        {
            var mgs = (groups ?? new List<string>()).Select(x => new MappingGroup { GroupKey = x, NormalizedKey = NormalizeSegmentGroupKey(x), DisplayKey = x }).ToList();
            return SuggestGroupMappings(mgs, pmsRows);
        }

        public static List<SuggestedMapping> SuggestGroupMappings(IList<MappingGroup> groups, IList<PmsRow> pmsRows)
        {
            var outList = new List<SuggestedMapping>();
            var pms = pmsRows ?? new List<PmsRow>();

            foreach (var g in groups ?? new List<MappingGroup>())
            {
                var best = FindBestSuggestion(g.GroupKey, pms);
                outList.Add(new SuggestedMapping
                {
                    GroupKey = g.GroupKey,
                    PmsClass = best?.Class ?? string.Empty,
                    PmsSegmentKey = best?.SegmentKey ?? string.Empty,
                    Score = best == null ? 0.0 : ComputeCandidateScore(g.GroupKey, best.SegmentKey)
                });
            }

            return outList.OrderByDescending(x => x.Score).ToList();
        }

        public static List<MappingSelection> ExpandGroupSelections(DataSet ds, IList<GroupSelection> selections)
        {
            var rules = ds?.Tables[TableRules];
            if (rules == null) return new List<MappingSelection>();

            var selectionMap = (selections ?? new List<GroupSelection>())
                .GroupBy(x => NormalizeSegmentGroupKey(x.GroupKey), StringComparer.OrdinalIgnoreCase)
                .ToDictionary(g => g.Key, g => g.First(), StringComparer.OrdinalIgnoreCase);

            var list = new List<MappingSelection>();
            foreach (DataRow r in rules.Rows)
            {
                var key = NormalizeSegmentGroupKey(SafeStr(r, "RevitSegmentKey"));
                if (!selectionMap.TryGetValue(key, out var sel)) continue;

                list.Add(new MappingSelection
                {
                    File = SafeStr(r, "File"),
                    PipeTypeName = SafeStr(r, "PipeTypeName"),
                    RuleIndex = SafeIntObj(r["RuleIndex"]),
                    SegmentId = SafeIntObj(r["SegmentId"]),
                    SegmentKey = key,
                    SelectedClass = sel.SelectedClass ?? string.Empty,
                    SelectedPmsSegment = sel.SelectedPmsSegment ?? string.Empty,
                    MappingSource = sel.SelectionSource ?? "group"
                });
            }

            return list;
        }

        public static RunResult RunCompare(DataSet ds,
            IList<PmsRow> pmsRows,
            IList<MappingSelection> mappings,
            CompareOptions options)
        {
            var mapTable = BuildMapTable();
            var revitSize = BuildRevitSizeTable();
            var pmsTable = BuildPmsTableSkeleton();
            var compare = BuildCompareTable();
            var errors = BuildErrorTable();
            var summary = BuildSummaryTable();

            foreach (var m in mappings ?? new List<MappingSelection>())
            {
                var r = mapTable.NewRow();
                r["File"] = m.File;
                r["PipeTypeName"] = m.PipeTypeName;
                r["RuleIndex"] = m.RuleIndex;
                r["SegmentId"] = m.SegmentId;
                r["RevitSegmentKey"] = m.SegmentKey;
                r["SelectedClass"] = m.SelectedClass;
                r["SelectedPmsSegment"] = m.SelectedPmsSegment;
                r["Source"] = m.MappingSource;
                mapTable.Rows.Add(r);
            }

            foreach (var p in pmsRows ?? new List<PmsRow>())
            {
                var r = pmsTable.NewRow();
                r["CLASS"] = p.Class;
                r["SEGMENT"] = p.SegmentKey;
                r["ND_mm"] = p.NdMm;
                r["ID_mm"] = p.IdMm;
                r["OD_mm"] = p.OdMm;
                pmsTable.Rows.Add(r);
            }

            // 비교 기본 행
            foreach (DataRow mr in mapTable.Rows)
            {
                var status = "CHECK";
                var detail = "비교 완료";
                var selectedSeg = SafeStr(mr, "SelectedPmsSegment");
                if (string.IsNullOrWhiteSpace(selectedSeg))
                {
                    status = "NO_MAPPING";
                    detail = "매핑 없음";
                }

                AddCompareRow(compare,
                    SafeStr(mr, "File"),
                    SafeStr(mr, "PipeTypeName"),
                    SafeIntObj(mr["RuleIndex"]),
                    SafeStr(mr, "RevitSegmentKey"),
                    SafeStr(mr, "SelectedClass"),
                    selectedSeg,
                    status,
                    detail);
            }

            var summaryRow = summary.NewRow();
            summaryRow["Item"] = "CompareCount";
            summaryRow["Value"] = compare.Rows.Count;
            summary.Rows.Add(summaryRow);

            var rr = new RunResult
            {
                MapTable = mapTable,
                RevitSizeTable = revitSize,
                PmsSizeTable = pmsTable,
                CompareTable = compare,
                ErrorTable = errors,
                SummaryTable = summary
            };

            rr.Summary["mapCount"] = mapTable.Rows.Count;
            rr.Summary["compareCount"] = compare.Rows.Count;
            rr.Summary["noMappingCount"] = compare.Rows.Cast<DataRow>().Count(x => SafeStr(x, "Status") == "NO_MAPPING");
            return rr;
        }

        public static List<Dictionary<string, object>> BuildClassCheckRows(DataTable mapTable)
        {
            return (mapTable?.Rows.Cast<DataRow>() ?? Enumerable.Empty<DataRow>())
                .Select(r => new Dictionary<string, object>
                {
                    ["File"] = SafeStr(r, "File"),
                    ["PipeTypeName"] = SafeStr(r, "PipeTypeName"),
                    ["RevitSegmentKey"] = SafeStr(r, "RevitSegmentKey"),
                    ["SelectedClass"] = SafeStr(r, "SelectedClass"),
                    ["SelectedPmsSegment"] = SafeStr(r, "SelectedPmsSegment")
                })
                .ToList();
        }

        public static List<Dictionary<string, object>> BuildSizeCheckRows(DataTable compareTable)
        {
            return (compareTable?.Rows.Cast<DataRow>() ?? Enumerable.Empty<DataRow>())
                .Select(r => new Dictionary<string, object>
                {
                    ["File"] = SafeStr(r, "File"),
                    ["PipeTypeName"] = SafeStr(r, "PipeTypeName"),
                    ["RuleIndex"] = SafeIntObj(r["RuleIndex"]),
                    ["Status"] = SafeStr(r, "Status"),
                    ["Detail"] = SafeStr(r, "Detail")
                })
                .ToList();
        }

        public static List<Dictionary<string, object>> BuildRoutingClassRows(DataSet ds)
        {
            var routing = ds?.Tables[TableRouting];
            return (routing?.Rows.Cast<DataRow>() ?? Enumerable.Empty<DataRow>())
                .Select(r => new Dictionary<string, object>
                {
                    ["File"] = SafeStr(r, "File"),
                    ["PipeTypeName"] = SafeStr(r, "PipeTypeName"),
                    ["RuleType"] = SafeStr(r, "RuleType"),
                    ["TypeName"] = SafeStr(r, "TypeName")
                })
                .ToList();
        }

        private static DataSet BuildExtractDataSet(ExtractOptions opts)
        {
            var ds = new DataSet("SegmentPms");
            ds.Tables.Add(BuildMetaTable(opts));
            ds.Tables.Add(BuildFileTable());
            ds.Tables.Add(BuildRuleTable());
            ds.Tables.Add(BuildSizeTable());
            ds.Tables.Add(BuildRoutingTable());
            return ds;
        }

        private static DataTable BuildMetaTable(ExtractOptions opts)
        {
            var t = new DataTable(TableMeta);
            t.Columns.Add("NdRound", typeof(int));
            t.Columns.Add("Tolerance", typeof(double));
            t.Columns.Add("CreatedAt", typeof(string));
            t.Columns.Add("ToolVersion", typeof(string));
            var r = t.NewRow();
            r["NdRound"] = opts?.NdRound ?? 3;
            r["Tolerance"] = opts?.ToleranceMm ?? 0.01;
            r["CreatedAt"] = DateTime.Now.ToString("s", CultureInfo.InvariantCulture);
            r["ToolVersion"] = "SegmentPms 2.0";
            t.Rows.Add(r);
            return t;
        }

        private static DataTable BuildFileTable()
        {
            var t = new DataTable(TableFiles);
            t.Columns.Add("File", typeof(string));
            t.Columns.Add("FileName", typeof(string));
            t.Columns.Add("ExtractedAt", typeof(string));
            return t;
        }

        private static DataTable BuildRuleTable()
        {
            var t = new DataTable(TableRules);
            t.Columns.Add("File", typeof(string));
            t.Columns.Add("FileName", typeof(string));
            t.Columns.Add("PipeTypeName", typeof(string));
            t.Columns.Add("RuleIndex", typeof(int));
            t.Columns.Add("RevitSegmentKey", typeof(string));
            t.Columns.Add("SegmentId", typeof(int));
            return t;
        }

        private static DataTable BuildSizeTable()
        {
            var t = new DataTable(TableSizes);
            t.Columns.Add("File", typeof(string));
            t.Columns.Add("SegmentId", typeof(int));
            t.Columns.Add("SegmentKey", typeof(string));
            t.Columns.Add("ND_mm", typeof(double));
            t.Columns.Add("ID_mm", typeof(double));
            t.Columns.Add("OD_mm", typeof(double));
            return t;
        }

        private static DataTable BuildRoutingTable()
        {
            var t = new DataTable(TableRouting);
            t.Columns.Add("File", typeof(string));
            t.Columns.Add("PipeTypeName", typeof(string));
            t.Columns.Add("RuleGroup", typeof(string));
            t.Columns.Add("RuleIndex", typeof(int));
            t.Columns.Add("RuleType", typeof(string));
            t.Columns.Add("PartId", typeof(int));
            t.Columns.Add("PartName", typeof(string));
            t.Columns.Add("TypeName", typeof(string));
            return t;
        }

        private static DataTable BuildMapTable()
        {
            var t = new DataTable("Map");
            t.Columns.Add("File", typeof(string));
            t.Columns.Add("PipeTypeName", typeof(string));
            t.Columns.Add("RuleIndex", typeof(int));
            t.Columns.Add("SegmentId", typeof(int));
            t.Columns.Add("RevitSegmentKey", typeof(string));
            t.Columns.Add("SelectedClass", typeof(string));
            t.Columns.Add("SelectedPmsSegment", typeof(string));
            t.Columns.Add("Source", typeof(string));
            return t;
        }

        private static DataTable BuildRevitSizeTable()
        {
            var t = new DataTable("RevitSize");
            t.Columns.Add("File", typeof(string));
            t.Columns.Add("SegmentId", typeof(int));
            t.Columns.Add("SegmentKey", typeof(string));
            t.Columns.Add("ND_mm", typeof(double));
            t.Columns.Add("ID_mm", typeof(double));
            t.Columns.Add("OD_mm", typeof(double));
            return t;
        }

        private static DataTable BuildPmsTableSkeleton()
        {
            var t = new DataTable("PMS");
            t.Columns.Add("CLASS", typeof(string));
            t.Columns.Add("SEGMENT", typeof(string));
            t.Columns.Add("ND_mm", typeof(double));
            t.Columns.Add("ID_mm", typeof(double));
            t.Columns.Add("OD_mm", typeof(double));
            return t;
        }

        private static DataTable BuildCompareTable()
        {
            var t = new DataTable("Compare");
            t.Columns.Add("File", typeof(string));
            t.Columns.Add("PipeTypeName", typeof(string));
            t.Columns.Add("RuleIndex", typeof(int));
            t.Columns.Add("RevitSegmentKey", typeof(string));
            t.Columns.Add("PmsClass", typeof(string));
            t.Columns.Add("PmsSegment", typeof(string));
            t.Columns.Add("Status", typeof(string));
            t.Columns.Add("Detail", typeof(string));
            return t;
        }

        private static DataTable BuildErrorTable()
        {
            var t = new DataTable("Error");
            t.Columns.Add("File", typeof(string));
            t.Columns.Add("Phase", typeof(string));
            t.Columns.Add("Message", typeof(string));
            return t;
        }

        private static DataTable BuildSummaryTable()
        {
            var t = new DataTable("Summary");
            t.Columns.Add("Item", typeof(string));
            t.Columns.Add("Value", typeof(object));
            return t;
        }

        private static void AddCompareRow(DataTable t,
            string file,
            string pipeType,
            int ruleIndex,
            string revitSegment,
            string pmsClass,
            string pmsSegment,
            string status,
            string detail)
        {
            var r = t.NewRow();
            r["File"] = file ?? string.Empty;
            r["PipeTypeName"] = pipeType ?? string.Empty;
            r["RuleIndex"] = ruleIndex;
            r["RevitSegmentKey"] = revitSegment ?? string.Empty;
            r["PmsClass"] = pmsClass ?? string.Empty;
            r["PmsSegment"] = pmsSegment ?? string.Empty;
            r["Status"] = status ?? string.Empty;
            r["Detail"] = detail ?? string.Empty;
            t.Rows.Add(r);
        }

        private static PmsRow FindBestSuggestion(string segmentKey, IList<PmsRow> rows)
        {
            var norm = NormalizeSegmentGroupKey(segmentKey);
            return (rows ?? new List<PmsRow>())
                .OrderByDescending(r => ComputeCandidateScore(norm, r?.SegmentKey))
                .FirstOrDefault();
        }

        private static double ComputeCandidateScore(string revitKey, string pmsKey)
        {
            if (string.IsNullOrWhiteSpace(revitKey) || string.IsNullOrWhiteSpace(pmsKey)) return 0;
            var a = TokenizeSegment(revitKey);
            var b = TokenizeSegment(pmsKey);
            if (a.Count == 0 || b.Count == 0) return 0;

            var inter = a.Intersect(b, StringComparer.OrdinalIgnoreCase).Count();
            var uni = a.Union(b, StringComparer.OrdinalIgnoreCase).Count();
            return uni == 0 ? 0 : (double)inter / uni;
        }

        private static HashSet<string> TokenizeSegment(string text)
        {
            var outSet = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            if (string.IsNullOrWhiteSpace(text)) return outSet;
            foreach (var token in Regex.Split(NormalizeSegmentGroupKey(text), "[^A-Za-z0-9]+"))
            {
                var t = (token ?? string.Empty).Trim();
                if (t.Length > 0) outSet.Add(t);
            }
            return outSet;
        }

        private static string NormalizeSegmentGroupKey(string s)
        {
            if (string.IsNullOrWhiteSpace(s)) return string.Empty;
            var v = s.Trim().ToUpperInvariant();
            v = Regex.Replace(v, "\\s+", " ");
            return v;
        }

        private static string SafeFileName(string p)
        {
            if (string.IsNullOrWhiteSpace(p)) return string.Empty;
            try { return Path.GetFileName(p); } catch { return p; }
        }

        private static bool ValidatePmsHeader(IRow row)
        {
            if (row == null) return false;
            var h0 = NormalizeHeader(CellStr(row.GetCell(0)));
            var h1 = NormalizeHeader(CellStr(row.GetCell(1)));
            var h2 = NormalizeHeader(CellStr(row.GetCell(2)));
            var h3 = NormalizeHeader(CellStr(row.GetCell(3)));
            var h4 = NormalizeHeader(CellStr(row.GetCell(4)));
            return h0 == "CLASS" && h1 == "SEGMENT" && h2 == "ND_MM" && h3 == "ID_MM" && h4 == "OD_MM";
        }

        private static string NormalizeHeader(string s)
        {
            var v = (s ?? string.Empty).Trim().ToUpperInvariant();
            v = v.Replace(" ", "_");
            return v;
        }

        private static string CellStr(ICell cell)
        {
            if (cell == null) return string.Empty;
            try
            {
                switch (cell.CellType)
                {
                    case CellType.Numeric: return cell.NumericCellValue.ToString(CultureInfo.InvariantCulture);
                    case CellType.Boolean: return cell.BooleanCellValue ? "TRUE" : "FALSE";
                    case CellType.Formula: return cell.ToString();
                    default: return cell.ToString();
                }
            }
            catch
            {
                return string.Empty;
            }
        }

        private static double CellDbl(ICell cell)
        {
            if (cell == null) return 0;
            try
            {
                if (cell.CellType == CellType.Numeric) return cell.NumericCellValue;
                var s = cell.ToString();
                if (double.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out var d)) return d;
                if (double.TryParse(s, NumberStyles.Any, CultureInfo.CurrentCulture, out d)) return d;
            }
            catch
            {
            }

            return 0;
        }

        private static string SafeStr(DataRow row, string col)
        {
            try
            {
                if (row != null && row.Table.Columns.Contains(col)) return Convert.ToString(row[col]) ?? string.Empty;
            }
            catch
            {
            }

            return string.Empty;
        }

        private static int SafeIntObj(object o)
        {
            try
            {
                if (o == null || o == DBNull.Value) return 0;
                return Convert.ToInt32(o);
            }
            catch
            {
                return 0;
            }
        }
    }
}
