using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;

namespace KKY_Tool_Revit.Infrastructure
{
    public static class ExcelCore
    {
        public static string PickAndSaveXlsx(
            string title,
            DataTable table,
            string defaultFileName,
            bool autoFit = false,
            string progressKey = null,
            string exportKind = null)
        {
            if (table == null) throw new ArgumentNullException(nameof(table));

            var path = PickSavePath("Excel Workbook (*.xlsx)|*.xlsx", defaultFileName, title);
            if (string.IsNullOrWhiteSpace(path)) return "";

            SaveXlsx(path, string.IsNullOrWhiteSpace(table.TableName) ? title : table.TableName, table, autoFit, sheetKey: title, progressKey: progressKey, exportKind: exportKind);
            return path;
        }

        public static string PickAndSaveXlsxMulti(
            IList<KeyValuePair<string, DataTable>> sheets,
            string defaultFileName,
            bool autoFit = false,
            string progressKey = null)
        {
            if (sheets == null || sheets.Count == 0) throw new ArgumentException("Sheets is empty.", nameof(sheets));

            var path = PickSavePath("Excel Workbook (*.xlsx)|*.xlsx", defaultFileName, "엑셀 저장");
            if (string.IsNullOrWhiteSpace(path)) return "";

            SaveXlsxMulti(path, sheets, autoFit, progressKey);
            return path;
        }

        public static void SaveXlsx(
            string filePath,
            string sheetName,
            DataTable table,
            bool autoFit = false,
            string sheetKey = null,
            string progressKey = null,
            string exportKind = null)
        {
            if (string.IsNullOrWhiteSpace(filePath)) throw new ArgumentNullException(nameof(filePath));
            if (table == null) throw new ArgumentNullException(nameof(table));

            EnsureDir(filePath);
            if (ShouldEnsureNoDataRow(sheetName, sheetKey, exportKind))
            {
                EnsureNoDataRow(table);
            }

            using (IWorkbook wb = new XSSFWorkbook())
            {
                var safeSheet = NormalizeSheetName(sheetName ?? "Sheet1");
                var sheet = wb.CreateSheet(safeSheet);
                WriteTableToSheet(wb, sheet, safeSheet, table, sheetKey, autoFit, progressKey, exportKind);

                using (var fs = new FileStream(filePath, FileMode.Create, FileAccess.Write, FileShare.None))
                {
                    wb.Write(fs);
                }
            }
        }

        public static void SaveXlsxMulti(
            string filePath,
            IList<KeyValuePair<string, DataTable>> sheets,
            bool autoFit = false,
            string progressKey = null)
        {
            if (string.IsNullOrWhiteSpace(filePath)) throw new ArgumentNullException(nameof(filePath));
            if (sheets == null || sheets.Count == 0) throw new ArgumentException("Sheets is empty.", nameof(sheets));

            EnsureDir(filePath);

            using (IWorkbook wb = new XSSFWorkbook())
            {
                var usedNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

                for (var i = 0; i < sheets.Count; i++)
                {
                    var name = sheets[i].Key ?? $"Sheet{i + 1}";
                    var table = sheets[i].Value;
                    if (table == null) continue;

                    var safe = MakeUniqueSheetName(NormalizeSheetName(name), usedNames);
                    usedNames.Add(safe);

                    if (ShouldEnsureNoDataRow(safe, name, null))
                    {
                        EnsureNoDataRow(table);
                    }

                    var sheet = wb.CreateSheet(safe);
                    WriteTableToSheet(wb, sheet, safe, table, sheetKey: name, autoFit: autoFit, progressKey: progressKey, exportKind: null);
                }

                using (var fs = new FileStream(filePath, FileMode.Create, FileAccess.Write, FileShare.None))
                {
                    wb.Write(fs);
                }
            }
        }

        public static void SaveStyledSimple(
            string filePath,
            string title,
            DataTable table,
            string groupHeader,
            bool autoFit = false,
            string progressKey = null)
        {
            if (string.IsNullOrWhiteSpace(filePath)) throw new ArgumentNullException(nameof(filePath));
            if (table == null) throw new ArgumentNullException(nameof(table));

            EnsureDir(filePath);

            using (IWorkbook wb = new XSSFWorkbook())
            {
                var baseName = string.IsNullOrWhiteSpace(title)
                    ? (string.IsNullOrWhiteSpace(table.TableName) ? "Sheet1" : table.TableName)
                    : title;

                var safeSheet = NormalizeSheetName(baseName);
                var sh = wb.CreateSheet(safeSheet);

                WriteTableToSheet(wb, sh, safeSheet, table, sheetKey: title, autoFit: false, progressKey: progressKey, exportKind: null);

                if (!string.IsNullOrWhiteSpace(groupHeader))
                {
                    TryApplyGroupBanding(wb, sh, table, groupHeader);
                }

                ApplyStandardSheetStyle(wb, sh, headerRowIndex: 0, autoFilter: true, freezeTopRow: true, borderAll: true, autoFit: autoFit);

                using (var fs = new FileStream(filePath, FileMode.Create, FileAccess.Write, FileShare.None))
                {
                    wb.Write(fs);
                }
            }
        }

        private static void TryApplyGroupBanding(IWorkbook wb, ISheet sh, DataTable table, string groupHeader)
        {
            if (wb == null || sh == null || table == null) return;
            if (string.IsNullOrWhiteSpace(groupHeader)) return;
            if (table.Columns.Count == 0 || table.Rows.Count == 0) return;

            var groupCol = -1;
            for (var c = 0; c < table.Columns.Count; c++)
            {
                if (string.Equals(table.Columns[c].ColumnName, groupHeader, StringComparison.OrdinalIgnoreCase))
                {
                    groupCol = c;
                    break;
                }
            }
            if (groupCol < 0) return;

            var cache = new Dictionary<int, ICellStyle>();
            string lastKey = null;
            var band = false;

            for (var r = 0; r < table.Rows.Count; r++)
            {
                var v = table.Rows[r][groupCol];
                var keyText = (v == null || v == DBNull.Value) ? "" : v.ToString();

                if (lastKey == null)
                {
                    lastKey = keyText;
                }
                else if (!string.Equals(lastKey, keyText, StringComparison.Ordinal))
                {
                    band = !band;
                    lastKey = keyText;
                }

                if (!band) continue;

                var row = sh.GetRow(r + 1);
                if (row == null) continue;

                var lastCol = table.Columns.Count - 1;
                for (var c = 0; c <= lastCol; c++)
                {
                    var cell = row.GetCell(c);
                    if (cell == null) continue;

                    var baseStyle = cell.CellStyle;
                    var styleKey = baseStyle == null ? -1 : (int)baseStyle.Index;

                    if (!cache.TryGetValue(styleKey, out var st))
                    {
                        st = wb.CreateCellStyle();
                        if (baseStyle != null) st.CloneStyleFrom(baseStyle);
                        st.FillForegroundColor = IndexedColors.Grey25Percent.Index;
                        st.FillPattern = FillPattern.SolidForeground;
                        cache[styleKey] = st;
                    }

                    cell.CellStyle = st;
                }
            }
        }

        private static readonly HashSet<string> ReviewExportKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "guid", "familylink", "paramprop", "pms", "sharedparambatch", "connector"
        };

        private static bool ShouldEnsureNoDataRow(string sheetName, string sheetKey, string exportKind)
        {
            var key = NormalizeExportPolicyKey(exportKind, sheetKey, sheetName);
            if (string.IsNullOrWhiteSpace(key)) return false;

            if (key.Equals("points", StringComparison.OrdinalIgnoreCase) ||
                key.Equals("export", StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }

            return ReviewExportKeys.Contains(key);
        }

        private static string NormalizeExportPolicyKey(string exportKind, string sheetKey, string sheetName)
        {
            string raw;
            if (!string.IsNullOrWhiteSpace(exportKind)) raw = exportKind;
            else if (!string.IsNullOrWhiteSpace(sheetKey)) raw = sheetKey;
            else raw = sheetName ?? "";

            if (string.IsNullOrWhiteSpace(raw)) return "";
            var s = raw.Trim().ToLowerInvariant();

            if (s.Contains("point")) return "points";
            if (s == "export") return "export";
            if (s.Contains("guid")) return "guid";
            if (s.Contains("familylink") || s.Contains("family link")) return "familylink";
            if (s.Contains("param")) return "paramprop";
            if (s.Contains("pms") || s.Contains("segment")) return "pms";
            if (s.Contains("sharedparambatch")) return "sharedparambatch";
            if (s.Contains("connector")) return "connector";

            return s;
        }

        public static void EnsureNoDataRow(DataTable table, string message = "오류가 없습니다.")
        {
            if (table == null) return;

            if (table.Columns.Count == 0)
            {
                table.Columns.Add("Message", typeof(string));
            }

            if (table.Rows.Count > 0) return;

            var finalMessage = string.IsNullOrWhiteSpace(message) ? "오류가 없습니다." : message;
            var row = table.NewRow();
            row[0] = finalMessage;
            table.Rows.Add(row);
        }

        public static void EnsureMessageRow(DataTable table, string message = "오류가 없습니다.")
        {
            EnsureNoDataRow(table, message);
        }

        public static void SaveCsv(string filePath, DataTable table)
        {
            if (string.IsNullOrWhiteSpace(filePath)) throw new ArgumentNullException(nameof(filePath));
            if (table == null) throw new ArgumentNullException(nameof(table));
            EnsureDir(filePath);

            using (var sw = new StreamWriter(filePath, false, new UTF8Encoding(encoderShouldEmitUTF8Identifier: true)))
            {
                for (var c = 0; c < table.Columns.Count; c++)
                {
                    if (c > 0) sw.Write(",");
                    sw.Write(EscapeCsv(table.Columns[c].ColumnName));
                }
                sw.WriteLine();

                for (var r = 0; r < table.Rows.Count; r++)
                {
                    var dr = table.Rows[r];
                    for (var c = 0; c < table.Columns.Count; c++)
                    {
                        if (c > 0) sw.Write(",");
                        var v = dr[c];
                        var s = v == null || v == DBNull.Value ? "" : v.ToString();
                        sw.Write(EscapeCsv(s));
                    }
                    sw.WriteLine();
                }
            }
        }

        private static void WriteTableToSheet(
            IWorkbook wb,
            ISheet sheet,
            string sheetName,
            DataTable table,
            string sheetKey,
            bool autoFit,
            string progressKey,
            string exportKind)
        {
            var colCount = table.Columns.Count;
            if (colCount == 0) return;

            var headerRow = sheet.CreateRow(0);
            var isConnector = string.Equals(exportKind, "connector", StringComparison.OrdinalIgnoreCase);
            var headerStyle = isConnector ? ExcelStyleHelper.GetHeaderStyleNoWrap(wb) : ExcelStyleHelper.GetHeaderStyle(wb);
            if (isConnector)
            {
                headerRow.Height = -1;
            }

            for (var c = 0; c < colCount; c++)
            {
                var cell = headerRow.CreateCell(c);
                cell.SetCellValue(table.Columns[c].ColumnName);
                cell.CellStyle = headerStyle;
            }

            sheet.CreateFreezePane(0, 1);

            var total = table.Rows.Count;
            for (var r = 0; r < total; r++)
            {
                var dr = table.Rows[r];
                var row = sheet.CreateRow(r + 1);
                if (isConnector)
                {
                    row.Height = -1;
                }

                for (var c = 0; c < colCount; c++)
                {
                    WriteCell(row, c, dr[c]);
                }

                var status = ExcelExportStyleRegistry.Resolve(sheetKey ?? sheetName, dr, table);
                if (status != ExcelStyleHelper.RowStatus.None)
                {
                    var style = isConnector ? ExcelStyleHelper.GetRowStyleNoWrap(wb, status) : ExcelStyleHelper.GetRowStyle(wb, status);
                    ExcelStyleHelper.ApplyStyleToRow(row, colCount, style);
                }

                if ((r % 200) == 0)
                {
                    TryReportProgress(progressKey, r, total, sheetName);
                }
            }

            if (isConnector && colCount > 0)
            {
                try
                {
                    var lastRowIndex = Math.Max(0, total);
                    var range = new CellRangeAddress(0, lastRowIndex, 0, colCount - 1);
                    sheet.SetAutoFilter(range);
                }
                catch
                {
                }
            }

            if (autoFit)
            {
                TryTrackAllColumnsForAutoSizing(sheet);
                for (var c = 0; c < colCount; c++)
                {
                    try
                    {
                        sheet.AutoSizeColumn(c);
                    }
                    catch
                    {
                    }
                }
            }
        }

        private static void WriteCell(IRow row, int colIndex, object value)
        {
            var cell = row.CreateCell(colIndex);

            if (value == null || value == DBNull.Value)
            {
                cell.SetCellValue("");
                return;
            }

            if (value is bool b)
            {
                cell.SetCellValue(b);
                return;
            }

            if (value is DateTime dt)
            {
                cell.SetCellValue(dt);
                return;
            }

            if (value is byte || value is short || value is int ||
                value is long || value is float || value is double ||
                value is decimal)
            {
                if (double.TryParse(value.ToString(), out var d))
                {
                    cell.SetCellValue(d);
                }
                else
                {
                    cell.SetCellValue(value.ToString());
                }
                return;
            }

            cell.SetCellValue(value.ToString());
        }

        private static string PickSavePath(string filter, string defaultFileName, string title)
        {
            using (var dlg = new SaveFileDialog())
            {
                dlg.Filter = filter;
                dlg.Title = string.IsNullOrWhiteSpace(title) ? "저장" : title;
                dlg.FileName = string.IsNullOrWhiteSpace(defaultFileName) ? "export.xlsx" : defaultFileName;
                dlg.RestoreDirectory = true;
                if (dlg.ShowDialog() != DialogResult.OK) return "";
                return dlg.FileName;
            }
        }

        private static void EnsureDir(string filePath)
        {
            var dir = Path.GetDirectoryName(filePath);
            if (!string.IsNullOrWhiteSpace(dir) && !Directory.Exists(dir))
            {
                Directory.CreateDirectory(dir);
            }
        }

        private static string NormalizeSheetName(string name)
        {
            var s = (name ?? "Sheet1").Trim();
            if (s.Length == 0) s = "Sheet1";

            var bad = new[] { ':', '\\', '/', '?', '*', '[', ']' };
            foreach (var ch in bad)
            {
                s = s.Replace(ch, '_');
            }

            if (s.Length > 31) s = s.Substring(0, 31);
            return s;
        }

        private static string MakeUniqueSheetName(string baseName, HashSet<string> used)
        {
            var s = baseName;
            var i = 1;
            while (used.Contains(s))
            {
                var suffix = $"({i})";
                var cut = Math.Min(31 - suffix.Length, baseName.Length);
                s = baseName.Substring(0, cut) + suffix;
                i += 1;
            }

            return s;
        }

        private static string EscapeCsv(string s)
        {
            if (s == null) return "";
            var needs = s.Contains(",") || s.Contains("\"") || s.Contains("\r") || s.Contains("\n");
            var t = s.Replace("\"", "\"\"");
            if (needs) return $"\"{t}\"";
            return t;
        }

        public static void ApplyStandardSheetStyle(
            IWorkbook wb,
            ISheet sh,
            int headerRowIndex = 0,
            bool autoFilter = true,
            bool freezeTopRow = true,
            bool borderAll = false,
            bool autoFit = false)
        {
            if (wb == null || sh == null) return;

            if (freezeTopRow)
            {
                try
                {
                    sh.CreateFreezePane(0, headerRowIndex + 1);
                }
                catch
                {
                }
            }

            if (autoFilter)
            {
                try
                {
                    var headerRow = sh.GetRow(headerRowIndex);
                    if (headerRow != null)
                    {
                        var lastCol = (int)headerRow.LastCellNum - 1;
                        if (lastCol >= 0)
                        {
                            var lastRow = Math.Max(headerRowIndex, sh.LastRowNum);
                            var range = new CellRangeAddress(headerRowIndex, lastRow, 0, lastCol);
                            TrySetAutoFilter(sh, range);
                        }
                    }
                }
                catch
                {
                }
            }

            if (borderAll)
            {
                TryApplyThinBorderToUsedRange(wb, sh);
            }

            if (autoFit)
            {
                TryTrackAllColumnsForAutoSizing(sh);
                var headerRow = sh.GetRow(headerRowIndex);
                var lastCol = headerRow == null ? -1 : (int)headerRow.LastCellNum - 1;
                if (lastCol >= 0)
                {
                    for (var c = 0; c <= lastCol; c++)
                    {
                        try
                        {
                            sh.AutoSizeColumn(c);
                        }
                        catch
                        {
                        }
                    }
                }
            }
        }

        public static void ApplyNumberFormatByHeader(
            IWorkbook wb,
            ISheet sh,
            int headerRowIndex,
            IEnumerable<string> headers,
            string numberFormat)
        {
            if (wb == null || sh == null || headers == null) return;

            var headerRow = sh.GetRow(headerRowIndex);
            if (headerRow == null) return;

            var headerSet = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var h in headers)
            {
                if (!string.IsNullOrWhiteSpace(h)) headerSet.Add(h.Trim());
            }
            if (headerSet.Count == 0) return;

            var targetCols = new List<int>();
            var fmt = new DataFormatter();
            var eval = wb.GetCreationHelper().CreateFormulaEvaluator();

            var lastCol = (int)headerRow.LastCellNum - 1;
            for (var c = 0; c <= lastCol; c++)
            {
                var cell = headerRow.GetCell(c);
                string text = "";
                try
                {
                    text = fmt.FormatCellValue(cell, eval).Trim();
                }
                catch
                {
                }

                if (headerSet.Contains(text)) targetCols.Add(c);
            }
            if (targetCols.Count == 0) return;

            var fmtIdx = wb.CreateDataFormat().GetFormat(string.IsNullOrWhiteSpace(numberFormat) ? "0.###############" : numberFormat);
            var styleCache = new Dictionary<int, ICellStyle>();

            for (var r = headerRowIndex + 1; r <= sh.LastRowNum; r++)
            {
                var row = sh.GetRow(r);
                if (row == null) continue;

                foreach (var c in targetCols)
                {
                    var cell = row.GetCell(c);
                    if (cell == null) continue;

                    if (cell.CellType == CellType.Numeric || cell.CellType == CellType.Formula)
                    {
                        var baseStyle = cell.CellStyle;
                        var key = baseStyle == null ? -1 : (int)baseStyle.Index;

                        if (!styleCache.TryGetValue(key, out var newStyle))
                        {
                            newStyle = wb.CreateCellStyle();
                            if (baseStyle != null) newStyle.CloneStyleFrom(baseStyle);
                            newStyle.DataFormat = fmtIdx;
                            styleCache[key] = newStyle;
                        }

                        cell.CellStyle = newStyle;
                    }
                }
            }
        }

        public static void ApplyResultFillByHeader(IWorkbook wb, ISheet sh, int headerRowIndex)
        {
            if (wb == null || sh == null) return;

            var headerRow = sh.GetRow(headerRowIndex);
            if (headerRow == null) return;

            var fmt = new DataFormatter();
            var eval = wb.GetCreationHelper().CreateFormulaEvaluator();

            var resultCol = -1;
            var lastCol = (int)headerRow.LastCellNum - 1;

            for (var c = 0; c <= lastCol; c++)
            {
                string h = "";
                try
                {
                    h = fmt.FormatCellValue(headerRow.GetCell(c), eval).Trim();
                }
                catch
                {
                }

                var norm = NormalizeHeader(h);
                if (norm == "result" || norm == "status")
                {
                    resultCol = c;
                    break;
                }
            }

            if (resultCol < 0) return;

            var warnCache = new Dictionary<int, ICellStyle>();
            var errCache = new Dictionary<int, ICellStyle>();

            for (var r = headerRowIndex + 1; r <= sh.LastRowNum; r++)
            {
                var row = sh.GetRow(r);
                if (row == null) continue;

                var cell = row.GetCell(resultCol);
                if (cell == null) continue;

                string text = "";
                try
                {
                    text = fmt.FormatCellValue(cell, eval);
                }
                catch
                {
                }

                var cls = ClassifyResult(text);
                if (cls == 0) continue;

                var baseStyle = cell.CellStyle;
                var key = baseStyle == null ? -1 : (int)baseStyle.Index;

                if (cls == 2)
                {
                    if (!errCache.TryGetValue(key, out var st))
                    {
                        st = wb.CreateCellStyle();
                        if (baseStyle != null) st.CloneStyleFrom(baseStyle);
                        st.FillForegroundColor = IndexedColors.Rose.Index;
                        st.FillPattern = FillPattern.SolidForeground;
                        errCache[key] = st;
                    }
                    cell.CellStyle = st;
                }
                else if (cls == 1)
                {
                    if (!warnCache.TryGetValue(key, out var st))
                    {
                        st = wb.CreateCellStyle();
                        if (baseStyle != null) st.CloneStyleFrom(baseStyle);
                        st.FillForegroundColor = IndexedColors.LightYellow.Index;
                        st.FillPattern = FillPattern.SolidForeground;
                        warnCache[key] = st;
                    }
                    cell.CellStyle = st;
                }
            }
        }

        public static void TryAutoFitWithExcel(string xlsxPath)
        {
            if (string.IsNullOrWhiteSpace(xlsxPath)) return;
            if (!File.Exists(xlsxPath)) return;

            object excelApp = null;
            object wbs = null;
            object wb = null;

            try
            {
                excelApp = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application"));
                if (excelApp == null) return;

                dynamic excel = excelApp;
                excel.DisplayAlerts = false;
                excel.Visible = false;

                wbs = excel.Workbooks;
                dynamic books = wbs;
                wb = books.Open(xlsxPath);

                object sheets = null;
                try
                {
                    sheets = ((dynamic)wb).Worksheets;
                    foreach (var ws in (dynamic)sheets)
                    {
                        try
                        {
                            ws.Cells.EntireColumn.AutoFit();
                        }
                        catch
                        {
                        }
                        finally
                        {
                            ReleaseCom(ws);
                        }
                    }
                }
                catch
                {
                }
                finally
                {
                    ReleaseCom(sheets);
                }

                try
                {
                    ((dynamic)wb).Save();
                }
                catch
                {
                }
            }
            catch
            {
            }
            finally
            {
                try
                {
                    if (wb != null) ((dynamic)wb).Close(SaveChanges: true);
                }
                catch
                {
                }

                try
                {
                    if (excelApp != null) ((dynamic)excelApp).Quit();
                }
                catch
                {
                }

                ReleaseCom(wb);
                ReleaseCom(wbs);
                ReleaseCom(excelApp);
            }
        }

        private static void ReleaseCom(object o)
        {
            try
            {
                if (o == null) return;
                if (Marshal.IsComObject(o))
                {
                    Marshal.FinalReleaseComObject(o);
                }
            }
            catch
            {
            }
        }

        private static void TryTrackAllColumnsForAutoSizing(ISheet sheet)
        {
            if (sheet == null) return;
            try
            {
                var mi = sheet.GetType().GetMethod("TrackAllColumnsForAutoSizing", Type.EmptyTypes);
                mi?.Invoke(sheet, null);
            }
            catch
            {
            }
        }

        private static void TrySetAutoFilter(ISheet sheet, CellRangeAddress range)
        {
            if (sheet == null || range == null) return;
            try
            {
                var mi = sheet.GetType().GetMethod("SetAutoFilter", new[] { typeof(CellRangeAddress) });
                mi?.Invoke(sheet, new object[] { range });
            }
            catch
            {
            }
        }

        private static void TryApplyThinBorderToUsedRange(IWorkbook wb, ISheet sh)
        {
            if (wb == null || sh == null) return;

            var maxCol = GetMaxUsedColumnIndex(sh);
            if (maxCol < 0) return;

            var cache = new Dictionary<int, ICellStyle>();

            for (var r = 0; r <= sh.LastRowNum; r++)
            {
                var row = sh.GetRow(r);
                if (row == null) continue;

                for (var c = 0; c <= maxCol; c++)
                {
                    var cell = row.GetCell(c);
                    if (cell == null)
                    {
                        cell = row.CreateCell(c);
                        cell.SetCellValue("");
                    }

                    var baseStyle = cell.CellStyle;
                    var key = baseStyle == null ? -1 : (int)baseStyle.Index;

                    if (!cache.TryGetValue(key, out var st))
                    {
                        st = wb.CreateCellStyle();
                        if (baseStyle != null) st.CloneStyleFrom(baseStyle);
                        st.BorderBottom = BorderStyle.Thin;
                        st.BorderTop = BorderStyle.Thin;
                        st.BorderLeft = BorderStyle.Thin;
                        st.BorderRight = BorderStyle.Thin;
                        cache[key] = st;
                    }

                    cell.CellStyle = st;
                }
            }
        }

        private static int GetMaxUsedColumnIndex(ISheet sh)
        {
            var maxCol = -1;
            for (var r = 0; r <= sh.LastRowNum; r++)
            {
                var row = sh.GetRow(r);
                if (row == null) continue;

                var lastCellNum = (int)row.LastCellNum;
                if (lastCellNum <= 0) continue;

                var lastIdx = lastCellNum - 1;
                if (lastIdx > maxCol) maxCol = lastIdx;
            }

            return maxCol;
        }

        private static string NormalizeHeader(string s)
        {
            if (string.IsNullOrWhiteSpace(s)) return "";
            return s.Trim().ToLowerInvariant().Replace(" ", "").Replace("_", "");
        }

        private static int ClassifyResult(string s)
        {
            if (string.IsNullOrWhiteSpace(s)) return 0;
            var t = s.Trim().ToLowerInvariant();

            if (t == "ok" || t == "pass" || t == "success") return 0;
            if (t.Contains("오류 없음") || t.Contains("정상") || t.Contains("이상 없음")) return 0;

            if (t.Contains("error") || t.Contains("fail") || t.Contains("mismatch")) return 2;
            if (t.Contains("실패") || t.Contains("오류") || t.Contains("불일치")) return 2;

            if (t.Contains("na") || t.Contains("n/a") || t.Contains("missing") || t.Contains("없음")) return 1;
            return 1;
        }

        private static void TryReportProgress(string progressKey, int current, int total, string sheetName)
        {
            if (string.IsNullOrWhiteSpace(progressKey)) return;
            try
            {
                var t = Type.GetType("KKY_Tool_Revit.UI.Hub.ExcelProgressReporter, " + typeof(ExcelCore).Assembly.FullName, throwOnError: false);
                if (t == null) return;

                var mi = t.GetMethod("Report", System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Static);
                if (mi == null) return;

                var percent = 0.0;
                if (total > 0) percent = ((double)current / total) * 100.0;
                mi.Invoke(null, new object[] { progressKey, percent, $"Exporting {sheetName}...", $"{current}/{total}" });
            }
            catch
            {
            }
        }
    }
}
