using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using Autodesk.Revit.UI;
using KKY_Tool_Revit.Infrastructure;
using KKY_Tool_Revit.Services;

namespace KKY_Tool_Revit.UI.Hub
{
    public sealed partial class UiBridgeExternalEvent
    {
        private static readonly Dictionary<string, double> ExportProgressWeights = new Dictionary<string, double>(StringComparer.OrdinalIgnoreCase)
        {
            ["COLLECT"] = 0.1,
            ["EXTRACT"] = 0.7,
            ["EXCEL"] = 0.05,
            ["EXCEL_INIT"] = 0.02,
            ["EXCEL_WRITE"] = 0.11,
            ["EXCEL_SAVE"] = 0.02,
            ["AUTOFIT"] = 0.0
        };

        private static readonly string[] ExportProgressOrder = { "COLLECT", "EXTRACT", "EXCEL", "EXCEL_INIT", "EXCEL_WRITE", "EXCEL_SAVE", "AUTOFIT" };
        private static DateTime _exportProgressLastSent = DateTime.MinValue;
        private static double _exportProgressLastPct;
        private static int _exportProgressLastRow;
        private static readonly object ExportProgressGate = new object();

        private static List<Dictionary<string, object>> _exportLastExportRows = new List<Dictionary<string, object>>();

        private void HandleExportBrowse()
        {
            using (var dlg = new System.Windows.Forms.FolderBrowserDialog())
            {
                var r = dlg.ShowDialog();
                if (r == System.Windows.Forms.DialogResult.OK)
                {
                    var files = Directory.GetFiles(dlg.SelectedPath, "*.rvt", SearchOption.AllDirectories);
                    SendToWeb("export:files", new { files });
                }
            }
        }

        private void HandleExportAddRvtFiles()
        {
            using (var dlg = new System.Windows.Forms.OpenFileDialog())
            {
                dlg.Filter = "Revit Project (*.rvt)|*.rvt";
                dlg.Multiselect = true;
                dlg.Title = "Export Points 대상 RVT 선택";
                dlg.RestoreDirectory = true;

                if (dlg.ShowDialog() != System.Windows.Forms.DialogResult.OK) return;

                var files = new List<string>();
                var dedup = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                foreach (var p in dlg.FileNames)
                {
                    if (string.IsNullOrWhiteSpace(p)) continue;
                    if (dedup.Add(p)) files.Add(p);
                }

                if (files.Count > 0)
                {
                    SendToWeb("export:rvt-files", new { files });
                }
            }
        }

        private void HandleExportPreview(UIApplication app, object payload)
        {
            ResetExportProgressState();
            try
            {
                var files = ExtractStringListLocal(payload, "files");
                ReportExportProgress("COLLECT", "파일 목록 준비 중", 0, files.Count, 0.0, true);

                var rows = TryCallExportPointsService(app, files);
                if (rows == null)
                {
                    ReportExportProgress("ERROR", "Export Points 서비스가 준비되지 않았습니다.", 0, 0, 0.0, true);
                    SendToWeb("revit:error", new { message = "Export Points 서비스가 준비되지 않았습니다." });
                    SendToWeb("export:previewed", new { rows = new List<Dictionary<string, object>>() });
                    return;
                }

                rows = rows.Select(AdaptExportRow).ToList();
                _exportLastExportRows = rows;
                SendToWeb("export:previewed", new { rows });
            }
            catch (Exception ex)
            {
                ReportExportProgress("ERROR", ex.Message, 0, 0, 0.0, true);
                SendToWeb("revit:error", new { message = "미리보기 실패: " + ex.Message });
                SendToWeb("export:previewed", new { rows = new List<Dictionary<string, object>>() });
            }
        }

        private void HandleExportSaveExcel(object payload)
        {
            ResetExportProgressState();
            try
            {
                var doAutoFit = ParseExcelMode(payload);
                var excelMode = ExtractExcelMode(payload, doAutoFit);
                var unit = ExtractUnit(payload);
                var rows = TryGetRowsFromPayload(payload);
                if (rows == null || rows.Count == 0) rows = _exportLastExportRows;
                if (rows == null) rows = new List<Dictionary<string, object>>();

                var total = rows.Count;
                ReportExportProgress("EXCEL", "엑셀 내보내기 준비 중", 0, total, 0.0, true);

                var dt = BuildExportDataTableFromRows(rows, unit, true);
                var todayToken = DateTime.Now.ToString("yyMMdd");
                var defaultName = $"{todayToken}_좌표 추출 결과.xlsx";
                var savePath = SaveExcelWithDialog(dt, defaultName, doAutoFit, excelMode);

                if (!string.IsNullOrEmpty(savePath))
                {
                    ReportExportProgress("DONE", "엑셀 내보내기 완료", total, total, 1.0, true);
                    SendToWeb("export:saved", new { path = savePath });
                }
                else
                {
                    ReportExportProgress("DONE", "엑셀 내보내기가 취소되었습니다.", total, total, 1.0, true);
                }
            }
            catch (Exception ex)
            {
                ReportExportProgress("ERROR", ex.Message, 0, 0, 0.0, true);
                SendToWeb("revit:error", new { message = "엑셀 내보내기 실패: " + ex.Message });
            }
        }

        private List<Dictionary<string, object>> TryCallExportPointsService(UIApplication app, List<string> files)
        {
            try
            {
                var direct = ExportPointsService.Run(app, files, HandleExportProgressFromService);
                return AnyToRows(direct);
            }
            catch
            {
            }

            var names = new[] { "KKY_Tool_Revit.Services.ExportPointsService", "Services.ExportPointsService" };
            foreach (var n in names)
            {
                var t = FindType(n, "ExportPointsService");
                if (t == null) continue;

                var m = t.GetMethod("Run", BindingFlags.Public | BindingFlags.Static | BindingFlags.Instance);
                if (m == null) continue;

                var inst = m.IsStatic ? null : Activator.CreateInstance(t);
                object[] args;
                var ps = m.GetParameters();

                if (ps != null && ps.Length >= 3)
                {
                    Delegate cb = null;
                    try
                    {
                        var cbMethod = typeof(UiBridgeExternalEvent).GetMethod(nameof(HandleExportProgressFromObject), BindingFlags.Instance | BindingFlags.NonPublic);
                        cb = Delegate.CreateDelegate(ps[2].ParameterType, this, cbMethod);
                    }
                    catch
                    {
                    }

                    args = cb != null ? new object[] { app, files, cb } : new object[] { app, files };
                }
                else
                {
                    args = new object[] { app, files };
                }

                var result = m.Invoke(inst, args);
                return AnyToRows(result);
            }

            return null;
        }

        private Dictionary<string, object> AdaptExportRow(Dictionary<string, object> r)
        {
            if (r == null) return new Dictionary<string, object>(StringComparer.Ordinal);

            var d = new Dictionary<string, object>(StringComparer.Ordinal)
            {
                ["File"] = r.ContainsKey("File") ? r["File"] : (r.ContainsKey("file") ? r["file"] : ""),
                ["ProjectPoint_E(mm)"] = FirstNonEmpty(r, new[] { "ProjectPoint_E(mm)", "ProjectE", "ProjectPoint_E", "ProjectPoint_E_ft" }),
                ["ProjectPoint_N(mm)"] = FirstNonEmpty(r, new[] { "ProjectPoint_N(mm)", "ProjectN", "ProjectPoint_N", "ProjectPoint_N_ft" }),
                ["ProjectPoint_Z(mm)"] = FirstNonEmpty(r, new[] { "ProjectPoint_Z(mm)", "ProjectZ", "ProjectPoint_Z", "ProjectPoint_Z_ft" }),
                ["SurveyPoint_E(mm)"] = FirstNonEmpty(r, new[] { "SurveyPoint_E(mm)", "SurveyE", "SurveyPoint_E", "SurveyPoint_E_ft" }),
                ["SurveyPoint_N(mm)"] = FirstNonEmpty(r, new[] { "SurveyPoint_N(mm)", "SurveyN", "SurveyPoint_N", "SurveyPoint_N_ft" }),
                ["SurveyPoint_Z(mm)"] = FirstNonEmpty(r, new[] { "SurveyPoint_Z(mm)", "SurveyZ", "SurveyPoint_Z", "SurveyPoint_Z_ft" }),
                ["TrueNorthAngle(deg)"] = FirstNonEmpty(r, new[] { "TrueNorthAngle(deg)", "TrueNorth", "TrueNorthAngle", "TrueNorthAngle_deg" })
            };

            return d;
        }

        private DataTable BuildExportDataTableFromRows(List<Dictionary<string, object>> rows, string unit, bool applyConversion)
        {
            var normalizedUnit = NormalizeUnit(unit);
            var suffix = "(ft)";
            if (normalizedUnit == "m") suffix = "(m)";
            else if (normalizedUnit == "mm") suffix = "(mm)";

            var dt = new DataTable("Export");
            var headers = new[]
            {
                "File",
                $"ProjectPoint_E{suffix}", $"ProjectPoint_N{suffix}", $"ProjectPoint_Z{suffix}",
                $"SurveyPoint_E{suffix}", $"SurveyPoint_N{suffix}", $"SurveyPoint_Z{suffix}",
                "TrueNorthAngle(deg)"
            };
            foreach (var h in headers) dt.Columns.Add(h);

            var total = rows?.Count ?? 0;
            var idx = 0;
            if (rows != null)
            {
                foreach (var r in rows)
                {
                    var dr = dt.NewRow();
                    dr[0] = SafeToString(r, "File");
                    dr[1] = FormatCoordForUnit(r, new[] { "ProjectPoint_E(mm)", "ProjectPoint_E(ft)", "ProjectPoint_E(m)", "ProjectE", "ProjectPoint_E" }, normalizedUnit, applyConversion);
                    dr[2] = FormatCoordForUnit(r, new[] { "ProjectPoint_N(mm)", "ProjectPoint_N(ft)", "ProjectPoint_N(m)", "ProjectN", "ProjectPoint_N" }, normalizedUnit, applyConversion);
                    dr[3] = FormatCoordForUnit(r, new[] { "ProjectPoint_Z(mm)", "ProjectPoint_Z(ft)", "ProjectPoint_Z(m)", "ProjectZ", "ProjectPoint_Z" }, normalizedUnit, applyConversion);
                    dr[4] = FormatCoordForUnit(r, new[] { "SurveyPoint_E(mm)", "SurveyPoint_E(ft)", "SurveyPoint_E(m)", "SurveyE", "SurveyPoint_E" }, normalizedUnit, applyConversion);
                    dr[5] = FormatCoordForUnit(r, new[] { "SurveyPoint_N(mm)", "SurveyPoint_N(ft)", "SurveyPoint_N(m)", "SurveyN", "SurveyPoint_N" }, normalizedUnit, applyConversion);
                    dr[6] = FormatCoordForUnit(r, new[] { "SurveyPoint_Z(mm)", "SurveyPoint_Z(ft)", "SurveyPoint_Z(m)", "SurveyZ", "SurveyPoint_Z" }, normalizedUnit, applyConversion);
                    dr[7] = FormatAngleValue(r, "TrueNorthAngle(deg)");
                    dt.Rows.Add(dr);

                    idx += 1;
                    var progress = total > 0 ? (double)idx / total : 1.0;
                    ReportExportProgress("EXCEL", "엑셀 데이터 구성", idx, total, progress, false);
                }
            }

            if (total == 0)
            {
                ReportExportProgress("EXCEL", "엑셀 데이터 구성", 0, 0, 0.0, false);
            }

            return dt;
        }

        private void HandleExportProgressFromService(ExportPointsService.ProgressInfo info)
        {
            if (info == null) return;
            ReportExportProgress(info.Phase, info.Message, info.Current, info.Total, info.PhaseProgress, false);
        }

        private void HandleExportProgressFromObject(object info)
        {
            if (info == null) return;
            try
            {
                var t = info.GetType();
                var phase = Convert.ToString(t.GetProperty("Phase")?.GetValue(info, null));
                var message = Convert.ToString(t.GetProperty("Message")?.GetValue(info, null));
                var current = ToIntSafe(t.GetProperty("Current")?.GetValue(info, null));
                var total = ToIntSafe(t.GetProperty("Total")?.GetValue(info, null));
                var phaseProgress = ToDoubleSafe(t.GetProperty("PhaseProgress")?.GetValue(info, null));
                ReportExportProgress(phase, message, current, total, phaseProgress, false);
            }
            catch
            {
            }
        }

        private static List<string> ExtractStringListLocal(object payload, string key)
        {
            var res = new List<string>();
            if (payload == null || string.IsNullOrEmpty(key)) return res;

            var v = GetProp(payload, key);
            if (v == null) return res;

            if (v is string s)
            {
                if (!string.IsNullOrEmpty(s)) res.Add(s);
                return res;
            }

            if (v is IEnumerable arr)
            {
                foreach (var o in arr)
                {
                    if (o == null) continue;
                    var txt = o.ToString();
                    if (!string.IsNullOrWhiteSpace(txt)) res.Add(txt);
                }
            }

            return res;
        }

        private static List<Dictionary<string, object>> AnyToRows(object any)
        {
            var result = new List<Dictionary<string, object>>();
            if (any == null) return result;

            if (any is List<Dictionary<string, object>> rows)
            {
                return rows;
            }

            if (any is DataTable dt)
            {
                return DataTableToRows(dt);
            }

            if (any is IEnumerable ie)
            {
                foreach (var item in ie)
                {
                    var d = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);
                    if (item is IDictionary dict)
                    {
                        foreach (var k in dict.Keys)
                        {
                            d[Convert.ToString(k)] = dict[k];
                        }
                    }
                    else if (item != null)
                    {
                        var t = item.GetType();
                        foreach (var p in t.GetProperties())
                        {
                            d[p.Name] = p.GetValue(item, null);
                        }
                    }

                    result.Add(d);
                }
            }

            return result;
        }

        private static string SafeToString(Dictionary<string, object> row, string col)
        {
            if (row == null) return string.Empty;
            if (row.TryGetValue(col, out var v) && v != null)
            {
                return Convert.ToString(v, CultureInfo.InvariantCulture);
            }

            return string.Empty;
        }

        private static double? SafeToDouble(Dictionary<string, object> row, IEnumerable<string> cols)
        {
            if (row == null || cols == null) return null;
            foreach (var k in cols)
            {
                var s = SafeToString(row, k);
                if (double.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out var d))
                {
                    return d;
                }
            }

            return null;
        }

        private static string FirstNonEmpty(Dictionary<string, object> row, IEnumerable<string> keys)
        {
            if (row == null || keys == null) return string.Empty;
            foreach (var k in keys)
            {
                var s = SafeToString(row, k);
                if (!string.IsNullOrEmpty(s)) return s;
            }

            return string.Empty;
        }

        private static string NormalizeUnit(string unit)
        {
            var u = (unit ?? "").Trim().ToLowerInvariant();
            if (u == "m" || u == "meter" || u == "meters") return "m";
            if (u == "mm" || u == "millimeter" || u == "millimeters") return "mm";
            return "ft";
        }

        private static double UnitFactor(string unit)
        {
            if (unit == "m") return 0.3048;
            if (unit == "mm") return 304.8;
            return 1.0;
        }

        private static string FormatCoordForUnit(Dictionary<string, object> row, IEnumerable<string> keys, string unit, bool applyConversion)
        {
            var val = SafeToDouble(row, keys);
            if (!val.HasValue) return string.Empty;

            var v = val.Value;
            if (applyConversion) v *= UnitFactor(unit);
            return v.ToString("0.####", CultureInfo.InvariantCulture);
        }

        private static string FormatAngleValue(Dictionary<string, object> row, string key)
        {
            var s = SafeToString(row, key);
            if (double.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out var d))
            {
                return d.ToString("0.###", CultureInfo.InvariantCulture);
            }

            return s;
        }

        private static string ExtractUnit(object payload)
        {
            if (payload == null) return "ft";
            var v = GetProp(payload, "unit");
            if (v != null)
            {
                return NormalizeUnit(Convert.ToString(v, CultureInfo.InvariantCulture));
            }

            return "ft";
        }

        private static string ExtractExcelMode(object payload, bool doAutoFit)
        {
            if (payload != null)
            {
                var v = GetProp(payload, "excelMode");
                if (v != null)
                {
                    var mode = Convert.ToString(v, CultureInfo.InvariantCulture);
                    if (!string.IsNullOrWhiteSpace(mode)) return mode.Trim().ToLowerInvariant();
                }
            }

            return doAutoFit ? "normal" : "fast";
        }

        private static List<Dictionary<string, object>> DataTableToRows(DataTable dt)
        {
            var list = new List<Dictionary<string, object>>();
            if (dt == null) return list;

            foreach (DataRow r in dt.Rows)
            {
                var d = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);
                foreach (DataColumn c in dt.Columns)
                {
                    d[c.ColumnName] = r.IsNull(c) ? null : r[c];
                }

                list.Add(d);
            }

            return list;
        }

        private static Type FindType(string fullOrSimple, string simpleMatch = null)
        {
            var t = Type.GetType(fullOrSimple, false);
            if (t != null) return t;

            foreach (var asm in AppDomain.CurrentDomain.GetAssemblies())
            {
                try
                {
                    t = asm.GetType(fullOrSimple, false);
                    if (t != null) return t;

                    if (!string.IsNullOrEmpty(simpleMatch))
                    {
                        foreach (var ti in asm.GetTypes())
                        {
                            if (string.Equals(ti.Name, simpleMatch, StringComparison.OrdinalIgnoreCase))
                            {
                                return ti;
                            }
                        }
                    }
                }
                catch
                {
                }
            }

            return null;
        }

        private static List<Dictionary<string, object>> TryGetRowsFromPayload(object payload)
        {
            if (payload == null) return new List<Dictionary<string, object>>();

            var rows = GetProp(payload, "rows");
            if (rows != null) return AnyToRows(rows);

            var data = GetProp(payload, "data");
            if (data != null) return AnyToRows(data);

            if (payload is IEnumerable ie)
            {
                return AnyToRows(ie);
            }

            return new List<Dictionary<string, object>>();
        }

        private static string SaveExcelWithDialog(DataTable dt, string defaultName = "export.xlsx", bool doAutoFit = false, string excelMode = "fast")
        {
            if (dt == null || dt.Columns.Count == 0) return string.Empty;

            var dlg = new Microsoft.Win32.SaveFileDialog
            {
                Filter = "Excel (*.xlsx)|*.xlsx",
                FileName = defaultName
            };

            var ok = dlg.ShowDialog();
            if (ok != true) return string.Empty;

            var path = dlg.FileName;
            var totalRows = dt.Rows.Count;
            ReportExportProgress("EXCEL_INIT", "엑셀 워크북 준비", 0, totalRows, 0.0, true);

            try
            {
                if (dt.Rows.Count == 0)
                {
                    ExcelCore.EnsureNoDataRow(dt, "추출 결과가 없습니다.");
                    totalRows = dt.Rows.Count;
                }

                ReportExportProgress("EXCEL_WRITE", "엑셀 데이터 작성", totalRows, totalRows, 1.0, true);
                ReportExportProgress("EXCEL_SAVE", "엑셀 파일 내보내기", totalRows, totalRows, 1.0, true);

                ExcelCore.SaveXlsx(path, "Export", dt, doAutoFit, sheetKey: "Export", progressKey: "export:progress", exportKind: "points");
                ExcelExportStyleRegistry.ApplyStylesForKey("points", path, autoFit: doAutoFit, excelMode: excelMode);

                var autoFitMessage = doAutoFit ? "AutoFit 적용" : "빠른 모드: AutoFit 생략";
                ReportExportProgress("AUTOFIT", autoFitMessage, totalRows, totalRows, 1.0, true);
                ReportExportProgress("DONE", "엑셀 내보내기 완료", totalRows, totalRows, 1.0, true);
                return path;
            }
            catch (Exception ex)
            {
                SendToWeb("host:error", new { message = "엑셀 내보내기 실패: " + ex.Message });
                ReportExportProgress("ERROR", ex.Message, 0, totalRows, 0.0, true);
                return string.Empty;
            }
        }

        private static int ToIntSafe(object obj)
        {
            if (obj == null) return 0;
            try { return Convert.ToInt32(obj); } catch { return 0; }
        }

        private static double ToDoubleSafe(object obj)
        {
            if (obj == null) return 0.0;
            try { return Convert.ToDouble(obj); } catch { return 0.0; }
        }

        private static void ResetExportProgressState()
        {
            lock (ExportProgressGate)
            {
                _exportProgressLastSent = DateTime.MinValue;
                _exportProgressLastPct = 0.0;
                _exportProgressLastRow = 0;
            }
        }

        private static void ReportExportProgress(string phase, string message, int current, int total, double phaseProgress, bool force = false)
        {
            var normalized = NormalizeExportPhase(phase);
            var pctToSend = 0.0;
            var shouldSend = false;
            var now = DateTime.UtcNow;

            lock (ExportProgressGate)
            {
                var computed = ComputeExportPercent(normalized, current, total, phaseProgress, _exportProgressLastPct);
                var elapsed = (now - _exportProgressLastSent).TotalMilliseconds;
                var delta = Math.Abs(computed - _exportProgressLastPct);
                var deltaRows = Math.Abs(current - _exportProgressLastRow);
                var important = normalized == "DONE" || normalized == "ERROR";

                if (force || important || elapsed >= 200.0 || delta >= 1.0 || deltaRows >= 200)
                {
                    _exportProgressLastSent = now;
                    _exportProgressLastPct = Math.Max(_exportProgressLastPct, computed);
                    _exportProgressLastRow = current;
                    pctToSend = _exportProgressLastPct;
                    shouldSend = true;
                }
            }

            if (!shouldSend) return;

            SendToWeb("export:progress", new
            {
                phase = normalized,
                message,
                current,
                total,
                phaseProgress = Clamp01(phaseProgress),
                percent = pctToSend
            });
        }

        private static double ComputeExportPercent(string phase, int current, int total, double phaseProgress, double lastPct)
        {
            if (phase == "DONE") return 100.0;
            if (phase == "ERROR") return lastPct;

            var completed = 0.0;
            var found = false;
            foreach (var key in ExportProgressOrder)
            {
                if (string.Equals(key, phase, StringComparison.OrdinalIgnoreCase))
                {
                    found = true;
                    break;
                }

                if (ExportProgressWeights.ContainsKey(key)) completed += ExportProgressWeights[key];
            }

            double weight;
            if (ExportProgressWeights.ContainsKey(phase))
            {
                weight = ExportProgressWeights[phase];
            }
            else if (!found)
            {
                weight = 1.0;
                completed = 0.0;
            }
            else
            {
                weight = 0.0;
            }

            var ratio = 0.0;
            if (total > 0) ratio = Math.Max(0.0, Math.Min(1.0, (double)current / total));
            ratio = Math.Max(ratio, Clamp01(phaseProgress));

            var pct = (completed + weight * ratio) * 100.0;
            if (pct < lastPct) return lastPct;
            if (pct > 100.0) return 100.0;
            return pct;
        }

        private static string NormalizeExportPhase(string phase)
        {
            var p = (phase ?? string.Empty).Trim().ToUpperInvariant();
            if (string.IsNullOrEmpty(p)) return "EXTRACT";
            return p;
        }

        private static double Clamp01(double v)
        {
            if (double.IsNaN(v) || double.IsInfinity(v)) return 0.0;
            if (v < 0.0) return 0.0;
            if (v > 1.0) return 1.0;
            return v;
        }
    }
}
