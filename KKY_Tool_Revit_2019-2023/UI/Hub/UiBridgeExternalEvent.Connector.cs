using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.Linq;
using Autodesk.Revit.UI;
using KKY_Tool_Revit.Infrastructure;

namespace KKY_Tool_Revit.UI.Hub
{
    public sealed partial class UiBridgeExternalEvent
    {
        private List<Dictionary<string, object>> lastConnRows;
        private List<Dictionary<string, object>> _connectorTotalRows;
        private List<Dictionary<string, object>> _connectorDetailRows;
        private List<string> _connectorExtraParams;
        private string _connectorTargetFilter = string.Empty;
        private bool _connectorExcludeEndDummy;
        private string _connectorUiUnit = "inch";

        private void LogDebug(string message)
        {
            try { SendToWeb("host:log", new { message = $"[{DateTime.Now:HH:mm:ss}] {message}" }); } catch { }
        }

        private void LogError(string message)
        {
            try { SendToWeb("host:error", new { message = $"[{DateTime.Now:HH:mm:ss}] {message}" }); } catch { }
        }

        private static void ReportConnectorProgress(double pct, string text)
        {
            SendToWeb("connector:progress", new { pct, text });
        }

        private static string SafePayloadSnapshot(object payload)
        {
            if (payload == null) return "(null)";
            try
            {
                if (payload is IDictionary<string, object> d)
                    return "{" + string.Join(", ", d.Select(kv => kv.Key + "=" + (kv.Value?.ToString() ?? "(null)"))) + "}";
                return payload.ToString();
            }
            catch { return "(payload)"; }
        }

        private void HandleConnectorRun(UIApplication app, object payload)
        {
            try
            {
                LogDebug("[connector] HandleConnectorRun 진입");
                LogDebug("[connector] payload=" + SafePayloadSnapshot(payload));

                var tol = TryToDouble(GetProp(payload, "tol"), 1.0);
                var unit = TryToString(GetProp(payload, "unit"), "inch");
                var param = TryToString(GetProp(payload, "param"), "Comments");
                _connectorExtraParams = ParseExtraParams(TryToString(GetProp(payload, "extraParams"), string.Empty));
                _connectorTargetFilter = TryToString(GetProp(payload, "targetFilter"), string.Empty);
                _connectorExcludeEndDummy = TryToBool(GetProp(payload, "excludeEndDummy"), false);
                _connectorUiUnit = NormalizeUiUnit(unit);

                var rows = Services.ConnectorDiagnosticsService.Run(app, tol, unit, param, _connectorExtraParams, _connectorTargetFilter, _connectorExcludeEndDummy, ReportConnectorProgress)
                           ?? new List<Dictionary<string, object>>();

                var totalRows = BuildTotalRows(rows);
                var filteredRows = totalRows.Where(ShouldIncludeRow).ToList();
                var mismatchAll = filteredRows.Where(IsMismatchRow).ToList();
                var nearAll = filteredRows.Where(IsNearConnection).ToList();
                const int previewLimit = 150;
                var previewRows = filteredRows.Take(previewLimit).ToList();

                _connectorTotalRows = filteredRows;
                _connectorDetailRows = rows;
                lastConnRows = filteredRows;

                SendToWeb("connector:loaded", new
                {
                    rows = previewRows,
                    total = filteredRows.Count,
                    previewCount = previewRows.Count,
                    mismatchCount = mismatchAll.Count,
                    nearCount = nearAll.Count,
                    hasMore = filteredRows.Count > previewLimit
                });

                SendToWeb("connector:done", new { ok = true, total = filteredRows.Count, mismatchCount = mismatchAll.Count, nearCount = nearAll.Count });
            }
            catch (Exception ex)
            {
                LogError("[connector] 실행 실패: " + ex.Message);
                SendToWeb("connector:done", new { ok = false, message = ex.Message });
            }
        }

        private void HandleConnectorSaveExcel(UIApplication app, object payload)
        {
            try
            {
                var mode = ParseExcelMode(payload);
                var autoFit = TryToBool(GetProp(payload, "autoFit"), false);

                var source = _connectorTotalRows ?? lastConnRows ?? _connectorDetailRows;
                if (source == null || source.Count == 0)
                {
                    SendToWeb("host:warn", new { message = "저장할 커넥터 결과가 없습니다." });
                    return;
                }

                var dt = ToDataTable(source);
                ExcelCore.EnsureNoDataRow(dt, "커넥터 결과가 없습니다.");

                var dlg = new Microsoft.Win32.SaveFileDialog
                {
                    Filter = "Excel (*.xlsx)|*.xlsx",
                    FileName = $"connector_diagnostics_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx"
                };

                if (dlg.ShowDialog() != true) return;

                ExcelCore.SaveXlsx(dlg.FileName, "Connector", dt, autoFit, sheetKey: "Connector", progressKey: "connector:progress", exportKind: "connector");
                ExcelExportStyleRegistry.ApplyStylesForKey("connector", dlg.FileName, autoFit, mode);

                SendToWeb("connector:saved", new { ok = true, path = dlg.FileName });
            }
            catch (Exception ex)
            {
                SendToWeb("connector:saved", new { ok = false, message = ex.Message });
            }
        }

        private static List<string> ParseExtraParams(string raw)
        {
            if (string.IsNullOrWhiteSpace(raw)) return new List<string>();
            return raw.Split(new[] { ',', ';', '|', '\n', '\r', '\t' }, StringSplitOptions.RemoveEmptyEntries)
                      .Select(x => x.Trim())
                      .Where(x => !string.IsNullOrWhiteSpace(x))
                      .Distinct(StringComparer.OrdinalIgnoreCase)
                      .ToList();
        }

        private static string NormalizeUiUnit(string unit)
        {
            var u = (unit ?? "inch").Trim().ToLowerInvariant();
            if (u == "in" || u == "inch" || u == "inches") return "inch";
            if (u == "ft" || u == "feet") return "ft";
            if (u == "mm") return "mm";
            if (u == "m") return "m";
            return "inch";
        }

        private static bool ShouldIncludeRow(Dictionary<string, object> row)
        {
            if (row == null) return false;
            var id1 = TryToString(Get(row, "Id1"), string.Empty);
            return !string.IsNullOrWhiteSpace(id1);
        }

        private static bool IsMismatchRow(Dictionary<string, object> row)
        {
            var status = TryToString(Get(row, "Status"), string.Empty);
            var cmp = TryToString(Get(row, "ParamCompare"), string.Empty);
            return status.Equals("ERROR", StringComparison.OrdinalIgnoreCase)
                   || cmp.Equals("Mismatch", StringComparison.OrdinalIgnoreCase);
        }

        private static bool IsNearConnection(Dictionary<string, object> row)
        {
            var connType = TryToString(Get(row, "ConnectionType"), string.Empty);
            return connType.Equals("PROXIMITY", StringComparison.OrdinalIgnoreCase);
        }

        private static List<Dictionary<string, object>> BuildTotalRows(List<Dictionary<string, object>> rows)
        {
            return rows ?? new List<Dictionary<string, object>>();
        }

        private static DataTable ToDataTable(IEnumerable<Dictionary<string, object>> rows)
        {
            var dt = new DataTable("Connector");
            var list = rows?.ToList() ?? new List<Dictionary<string, object>>();
            var columns = new List<string>();
            var set = new HashSet<string>(StringComparer.Ordinal);
            foreach (var r in list)
            {
                if (r == null) continue;
                foreach (var k in r.Keys)
                {
                    if (set.Add(k)) columns.Add(k);
                }
            }

            foreach (var c in columns) dt.Columns.Add(c, typeof(string));
            foreach (var r in list)
            {
                var dr = dt.NewRow();
                foreach (var c in columns) dr[c] = ConvertToText(Get(r, c));
                dt.Rows.Add(dr);
            }

            return dt;
        }

        private static object Get(Dictionary<string, object> row, string key)
        {
            if (row == null || key == null) return null;
            return row.TryGetValue(key, out var v) ? v : null;
        }

        private static string ConvertToText(object v)
        {
            if (v == null) return string.Empty;
            if (v is double d) return d.ToString("0.########", CultureInfo.InvariantCulture);
            return v.ToString() ?? string.Empty;
        }

        private static string TryToString(object value, string fallback)
        {
            try
            {
                var s = value?.ToString();
                return string.IsNullOrWhiteSpace(s) ? fallback : s;
            }
            catch { return fallback; }
        }

        private static double TryToDouble(object value, double fallback)
        {
            if (value == null) return fallback;
            try
            {
                if (value is double d) return d;
                if (double.TryParse(value.ToString(), NumberStyles.Any, CultureInfo.InvariantCulture, out var v)) return v;
                if (double.TryParse(value.ToString(), NumberStyles.Any, CultureInfo.CurrentCulture, out v)) return v;
            }
            catch { }
            return fallback;
        }

        private static bool TryToBool(object value, bool fallback)
        {
            if (value == null) return fallback;
            try
            {
                if (value is bool b) return b;
                if (bool.TryParse(value.ToString(), out var bv)) return bv;
                if (int.TryParse(value.ToString(), out var iv)) return iv != 0;
            }
            catch { }
            return fallback;
        }
    }
}
