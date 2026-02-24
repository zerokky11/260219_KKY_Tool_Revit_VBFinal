using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace KKY_Tool_Revit.Infrastructure
{
    public static class ExcelExportStyleRegistry
    {
        public delegate ExcelStyleHelper.RowStatus RowStatusResolver(DataRow row, DataTable table);

        private static readonly Dictionary<string, RowStatusResolver> Resolvers = new Dictionary<string, RowStatusResolver>(StringComparer.OrdinalIgnoreCase);

        static ExcelExportStyleRegistry()
        {
            Register("connector", ResolveConnector);
            Register("guid", ResolveResultLike);
            Register("paramprop", ResolveParamProp);
            Register("points", ResolveResultLike);
            Register("pms", ResolveResultLike);
            Register("familylink", ResolveIssueLike);
            Register("sharedparambatch", ResolveSharedParamBatch);

            Register("connector diagnostics", ResolveConnector);
            Register("familylinkaudit", ResolveIssueLike);
            Register("family link audit", ResolveIssueLike);
            Register("pms vs segment size검토", ResolveResultLike);
            Register("pipe segment class검토", ResolveResultLike);
            Register("routing class검토", ResolveResultLike);
        }

        public static void Register(string key, RowStatusResolver resolver)
        {
            if (string.IsNullOrWhiteSpace(key) || resolver == null) return;
            Resolvers[key.Trim()] = resolver;
        }

        public static ExcelStyleHelper.RowStatus Resolve(string sheetNameOrKey, DataRow row, DataTable table)
        {
            if (row == null || table == null) return ExcelStyleHelper.RowStatus.None;

            var key = NormalizeKey(sheetNameOrKey);
            if (!string.IsNullOrWhiteSpace(key) && Resolvers.TryGetValue(key, out var fn))
            {
                return SafeResolve(fn, row, table);
            }

            if (!string.IsNullOrWhiteSpace(sheetNameOrKey) && Resolvers.TryGetValue(sheetNameOrKey.Trim(), out fn))
            {
                return SafeResolve(fn, row, table);
            }

            return ResolveGeneric(row, table);
        }

        public static DataTable FilterIssueRows(string styleKey, DataTable table)
        {
            if (table == null) return null;

            var filtered = table.Clone();
            foreach (DataRow row in table.Rows)
            {
                var status = Resolve(styleKey, row, table);
                if (status == ExcelStyleHelper.RowStatus.Warning || status == ExcelStyleHelper.RowStatus.Error)
                {
                    filtered.ImportRow(row);
                }
            }
            return filtered;
        }

        public static void ApplyStylesForKey(string styleKey, string xlsxPath, string sheetName = null, bool autoFit = true, string excelMode = "normal")
        {
            if (string.IsNullOrWhiteSpace(xlsxPath)) return;
            if (!File.Exists(xlsxPath)) return;

            using (var fs = new FileStream(xlsxPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                IWorkbook wb = new XSSFWorkbook(fs);

                var targetSheets = new List<ISheet>();
                if (!string.IsNullOrWhiteSpace(sheetName))
                {
                    var namedSheet = wb.GetSheet(sheetName);
                    if (namedSheet != null) targetSheets.Add(namedSheet);
                }

                if (targetSheets.Count == 0)
                {
                    for (var i = 0; i < wb.NumberOfSheets; i++)
                    {
                        var sh = wb.GetSheetAt(i);
                        if (sh != null) targetSheets.Add(sh);
                    }
                }

                foreach (var sh in targetSheets)
                {
                    ApplyStandardSheetStyle(wb, sh);
                    ApplyBorders(wb, sh);
                    ApplyKeySpecificFormats(styleKey, sh);
                }

                fs.Close();
                using (var outFs = new FileStream(xlsxPath, FileMode.Create, FileAccess.Write, FileShare.None))
                {
                    wb.Write(outFs);
                }
            }
        }

        private static void ApplyStandardSheetStyle(IWorkbook wb, ISheet sh)
        {
            if (wb == null || sh == null) return;

            var header = sh.GetRow(0);
            if (header == null) return;

            var headerStyle = ExcelStyleHelper.GetHeaderStyle(wb);
            if (headerStyle == null) return;

            var lastCol = header.LastCellNum - 1;
            if (lastCol < 0) return;

            for (var c = 0; c <= lastCol; c++)
            {
                var cell = header.GetCell(c) ?? header.CreateCell(c);
                if (cell.CellType == CellType.Blank) cell.SetCellValue(string.Empty);
                cell.CellStyle = headerStyle;
            }
        }

        private static void ApplyBorders(IWorkbook wb, ISheet sh)
        {
            if (wb == null || sh == null) return;
            var lastRow = sh.LastRowNum;
            if (lastRow < 0) return;

            for (var r = 0; r <= lastRow; r++)
            {
                var row = sh.GetRow(r);
                if (row == null) continue;

                var lastCol = row.LastCellNum - 1;
                if (lastCol < 0) continue;

                for (var c = 0; c <= lastCol; c++)
                {
                    var cell = row.GetCell(c) ?? row.CreateCell(c);
                    if (cell.CellType == CellType.Blank) cell.SetCellValue(string.Empty);

                    var src = cell.CellStyle;
                    var dst = wb.CreateCellStyle();
                    if (src != null) dst.CloneStyleFrom(src);
                    dst.BorderBottom = BorderStyle.Thin;
                    dst.BorderTop = BorderStyle.Thin;
                    dst.BorderLeft = BorderStyle.Thin;
                    dst.BorderRight = BorderStyle.Thin;
                    cell.CellStyle = dst;
                }
            }
        }

        private static void ApplyKeySpecificFormats(string styleKey, ISheet sh)
        {
            if (sh == null) return;
            var key = NormalizeKey(styleKey);
            if (string.IsNullOrWhiteSpace(key)) return;

            if (key == "pms")
            {
                return;
            }
        }

        private static ExcelStyleHelper.RowStatus SafeResolve(RowStatusResolver fn, DataRow row, DataTable table)
        {
            try
            {
                return fn(row, table);
            }
            catch
            {
                return ExcelStyleHelper.RowStatus.None;
            }
        }

        private static string NormalizeKey(string sheetNameOrKey)
        {
            if (string.IsNullOrWhiteSpace(sheetNameOrKey)) return string.Empty;
            var s = sheetNameOrKey.Trim().ToLowerInvariant();
            if (s.Contains("connector")) return "connector";
            if (s.Contains("guid")) return "guid";
            if (s.Contains("param")) return "paramprop";
            if (s.Contains("point")) return "points";
            if (s.Contains("sharedparambatch")) return "sharedparambatch";
            if (s.Contains("familylink") || s.Contains("family link")) return "familylink";
            if (s.Contains("pms") || s.Contains("segment")) return "pms";
            return s;
        }

        private static ExcelStyleHelper.RowStatus ResolveConnector(DataRow row, DataTable table)
        {
            var statusText = GetFirstExistingText(row, table, "Status", "Result", "검토결과");
            if (IsOkLike(statusText)) return ExcelStyleHelper.RowStatus.None;
            if (LooksError(statusText)) return ExcelStyleHelper.RowStatus.Error;
            return ExcelStyleHelper.RowStatus.Warning;
        }

        private static ExcelStyleHelper.RowStatus ResolveIssueLike(DataRow row, DataTable table)
        {
            var issue = GetFirstExistingText(row, table, "Issue", "Result", "Status");
            if (IsOkLike(issue)) return ExcelStyleHelper.RowStatus.None;
            if (LooksError(issue)) return ExcelStyleHelper.RowStatus.Error;
            return ExcelStyleHelper.RowStatus.Warning;
        }

        private static ExcelStyleHelper.RowStatus ResolveResultLike(DataRow row, DataTable table)
        {
            var result = GetFirstExistingText(row, table, "Result", "Res", "Status", "성공여부", "검토결과", "Class검토결과", "Class검토");
            if (IsOkLike(result)) return ExcelStyleHelper.RowStatus.None;
            if (LooksError(result)) return ExcelStyleHelper.RowStatus.Error;
            if (string.IsNullOrWhiteSpace(result)) return ExcelStyleHelper.RowStatus.Info;
            return ExcelStyleHelper.RowStatus.Warning;
        }

        private static ExcelStyleHelper.RowStatus ResolveParamProp(DataRow row, DataTable table)
        {
            var detail = GetFirstExistingText(row, table, "Detail", "Result", "Status", "메시지");
            if (string.IsNullOrWhiteSpace(detail)) return ExcelStyleHelper.RowStatus.None;
            if (IsOkLike(detail)) return ExcelStyleHelper.RowStatus.None;
            if (LooksError(detail)) return ExcelStyleHelper.RowStatus.Error;
            return ExcelStyleHelper.RowStatus.Warning;
        }

        private static ExcelStyleHelper.RowStatus ResolveSharedParamBatch(DataRow row, DataTable table)
        {
            var status = GetFirstExistingText(row, table, "성공여부", "Level", "Status", "Result", "메시지");
            if (IsOkLike(status)) return ExcelStyleHelper.RowStatus.None;
            if (LooksError(status)) return ExcelStyleHelper.RowStatus.Error;
            if (string.IsNullOrWhiteSpace(status)) return ExcelStyleHelper.RowStatus.None;
            return ExcelStyleHelper.RowStatus.Warning;
        }

        private static ExcelStyleHelper.RowStatus ResolveGeneric(DataRow row, DataTable table)
        {
            var txt = GetFirstExistingText(row, table, "Status", "Result", "Issue", "Error", "ErrorMessage", "Notes");
            if (string.IsNullOrWhiteSpace(txt)) return ExcelStyleHelper.RowStatus.None;
            if (IsOkLike(txt)) return ExcelStyleHelper.RowStatus.None;
            if (LooksError(txt)) return ExcelStyleHelper.RowStatus.Error;
            return ExcelStyleHelper.RowStatus.Warning;
        }

        private static string GetFirstExistingText(DataRow row, DataTable table, params string[] colNames)
        {
            foreach (var c in colNames)
            {
                var t = GetColText(row, table, c);
                if (!string.IsNullOrWhiteSpace(t)) return t;
            }
            return string.Empty;
        }

        private static string GetColText(DataRow row, DataTable table, string colName)
        {
            foreach (DataColumn col in table.Columns)
            {
                if (!string.Equals(col.ColumnName, colName, StringComparison.OrdinalIgnoreCase)) continue;
                var v = row[col];
                if (v == null || v is DBNull) return string.Empty;
                return v.ToString().Trim();
            }
            return string.Empty;
        }

        private static bool IsOkLike(string s)
        {
            if (string.IsNullOrWhiteSpace(s)) return false;
            var t = s.Trim().ToLowerInvariant();
            if (t == "ok" || t == "pass" || t == "success") return true;
            if (t.StartsWith("ok(") || t.StartsWith("ok[") || t.StartsWith("ok_")) return true;
            if (t.Contains("오류 없음") || t.Contains("정상") || t.Contains("이상 없음")) return true;
            return false;
        }

        private static bool LooksError(string s)
        {
            if (string.IsNullOrWhiteSpace(s)) return false;
            var t = s.Trim().ToLowerInvariant();
            if (t.Contains("mismatch") || t.Contains("불일치")) return true;
            if (t.Contains("error") || t.Contains("fail")) return true;
            if (t.Contains("실패") || t.Contains("오류")) return true;
            return false;
        }
    }
}
