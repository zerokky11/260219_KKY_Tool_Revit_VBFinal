using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using Autodesk.Revit.UI;
using KKY_Tool_Revit.Services;
using KKY_Tool_Revit.Infrastructure;

namespace KKY_Tool_Revit.UI.Hub
{
    public sealed partial class UiBridgeExternalEvent
    {
        private sealed class MultiCommonOptions
        {
            public string ExtraParams { get; set; } = string.Empty;
            public string TargetFilter { get; set; } = string.Empty;
            public bool ExcludeEndDummy { get; set; }
        }

        private sealed class MultiConnectorOptions
        {
            public bool Enabled { get; set; } = true;
            public double Tol { get; set; } = 1.0;
            public string Unit { get; set; } = "inch";
            public string Param { get; set; } = "Comments";
        }

        private sealed class MultiPmsOptions
        {
            public bool Enabled { get; set; }
            public int NdRound { get; set; } = 3;
            public double TolMm { get; set; } = 0.01;
            public bool ClassMatch { get; set; }
        }

        private sealed class MultiGuidOptions
        {
            public bool Enabled { get; set; }
            public bool IncludeFamily { get; set; }
            public bool IncludeAnnotation { get; set; }
        }

        private sealed class MultiRunRequest
        {
            public MultiCommonOptions Common { get; set; } = new MultiCommonOptions();
            public MultiConnectorOptions Connector { get; set; } = new MultiConnectorOptions();
            public MultiPmsOptions Pms { get; set; } = new MultiPmsOptions();
            public MultiGuidOptions Guid { get; set; } = new MultiGuidOptions();
            public List<string> RvtPaths { get; set; } = new List<string>();
        }

        private readonly List<Dictionary<string, object>> _multiConnectorRows = new List<Dictionary<string, object>>();
        private readonly List<Dictionary<string, object>> _multiPmsClassRows = new List<Dictionary<string, object>>();
        private readonly List<Dictionary<string, object>> _multiPmsSizeRows = new List<Dictionary<string, object>>();
        private readonly List<Dictionary<string, object>> _multiPmsRoutingRows = new List<Dictionary<string, object>>();
        private readonly List<Dictionary<string, object>> _multiGuidRows = new List<Dictionary<string, object>>();
        private readonly List<Dictionary<string, object>> _multiRunItems = new List<Dictionary<string, object>>();

        private void HandleCommonOptionsGet(UIApplication app, object payload)
        {
            try
            {
                var stored = HubCommonOptionsStorageService.Load();
                SendToWeb("commonoptions:loaded", new
                {
                    extraParamsText = stored?.ExtraParamsText ?? string.Empty,
                    targetFilterText = stored?.TargetFilterText ?? string.Empty,
                    excludeEndDummy = stored?.ExcludeEndDummy ?? false
                });
            }
            catch (Exception ex)
            {
                SendToWeb("commonoptions:loaded", new
                {
                    extraParamsText = string.Empty,
                    targetFilterText = string.Empty,
                    excludeEndDummy = false,
                    errorMessage = ex.Message
                });
            }
        }

        private void HandleCommonOptionsSave(UIApplication app, object payload)
        {
            var extraText = GetString(payload, "extraParamsText", string.Empty);
            var filterText = GetString(payload, "targetFilterText", string.Empty);
            var excludeEndDummy = GetBool(payload, "excludeEndDummy", false);

            var options = new HubCommonOptionsStorageService.HubCommonOptions
            {
                ExtraParamsText = extraText ?? string.Empty,
                TargetFilterText = filterText ?? string.Empty,
                ExcludeEndDummy = excludeEndDummy
            };

            var ok = HubCommonOptionsStorageService.Save(options);
            SendToWeb("commonoptions:saved", new { ok });
        }

        private void HandleMultiPickRvt(UIApplication app, object payload)
        {
            var files = new Microsoft.Win32.OpenFileDialog { Filter = "RVT (*.rvt)|*.rvt", Multiselect = true };
            var selected = files.ShowDialog() == true ? files.FileNames.ToList() : new List<string>();
            SendToWeb("hub:rvt-picked", new { paths = selected });
        }

        private void HandleMultiClear(UIApplication app, object payload)
        {
            _multiConnectorRows.Clear();
            _multiPmsClassRows.Clear();
            _multiPmsSizeRows.Clear();
            _multiPmsRoutingRows.Clear();
            _multiGuidRows.Clear();
            _multiRunItems.Clear();
            SendToWeb("multi:review-summary", BuildMultiSummaryPayload());
        }

        private MultiRunRequest ParseMultiRequest(object payload)
        {
            var req = new MultiRunRequest();
            req.RvtPaths = GetStringList(payload, "files").Where(x => !string.IsNullOrWhiteSpace(x)).Distinct(StringComparer.OrdinalIgnoreCase).ToList();

            req.Connector.Enabled = GetBool(payload, "runConnector", true);
            req.Connector.Tol = GetDouble(payload, "tol", 1.0);
            req.Connector.Unit = GetString(payload, "unit", "inch");
            req.Connector.Param = GetString(payload, "param", "Comments");

            req.Pms.Enabled = GetBool(payload, "runSegmentPms", false);
            req.Pms.NdRound = GetInt(payload, "ndRound", 3);
            req.Pms.TolMm = GetDouble(payload, "tolMm", 0.01);
            req.Pms.ClassMatch = GetBool(payload, "classMatch", false);

            req.Guid.Enabled = GetBool(payload, "runGuid", false);
            req.Guid.IncludeFamily = GetBool(payload, "includeFamily", false);
            req.Guid.IncludeAnnotation = GetBool(payload, "includeAnnotation", false);

            req.Common.ExtraParams = GetString(payload, "extraParamsText", string.Empty);
            req.Common.TargetFilter = GetString(payload, "targetFilterText", string.Empty);
            req.Common.ExcludeEndDummy = GetBool(payload, "excludeEndDummy", false);
            return req;
        }

        private void HandleMultiRun(UIApplication app, object payload)
        {
            var req = ParseMultiRequest(payload);
            if (req.RvtPaths.Count == 0)
            {
                SendToWeb("hub:multi-error", new { message = "검토할 RVT 파일이 없습니다." });
                return;
            }

            if (!req.Connector.Enabled && !req.Pms.Enabled && !req.Guid.Enabled)
            {
                SendToWeb("hub:multi-error", new { message = "선택된 기능이 없습니다." });
                return;
            }

            _multiConnectorRows.Clear();
            _multiPmsClassRows.Clear();
            _multiPmsSizeRows.Clear();
            _multiPmsRoutingRows.Clear();
            _multiGuidRows.Clear();
            _multiRunItems.Clear();

            for (var i = 0; i < req.RvtPaths.Count; i++)
            {
                var filePath = req.RvtPaths[i];
                var safe = Path.GetFileName(filePath);
                SendToWeb("hub:multi-progress", new { current = i + 1, total = req.RvtPaths.Count, message = $"처리 중: {safe}" });

                var started = DateTime.Now;

                try
                {
                    if (!File.Exists(filePath))
                    {
                        AppendRunItem(safe, "skipped", "파일을 찾을 수 없습니다.", "OPEN", started);
                        AppendErrorRow(_multiConnectorRows, safe, "파일을 찾을 수 없습니다.");
                        continue;
                    }

                    if (req.Connector.Enabled)
                    {
                        _multiConnectorRows.Add(new Dictionary<string, object>
                        {
                            ["File"] = safe,
                            ["Status"] = "CHECK",
                            ["Detail"] = "Connector check queued",
                            ["Tol"] = req.Connector.Tol,
                            ["Unit"] = req.Connector.Unit,
                            ["Param"] = req.Connector.Param
                        });
                    }

                    if (req.Pms.Enabled)
                    {
                        _multiPmsClassRows.Add(new Dictionary<string, object>
                        {
                            ["File"] = safe,
                            ["Status"] = "CHECK",
                            ["Detail"] = "PMS class compare queued"
                        });
                        _multiPmsSizeRows.Add(new Dictionary<string, object>
                        {
                            ["File"] = safe,
                            ["Status"] = "CHECK",
                            ["Detail"] = "PMS size compare queued"
                        });
                        _multiPmsRoutingRows.Add(new Dictionary<string, object>
                        {
                            ["File"] = safe,
                            ["Status"] = "CHECK",
                            ["Detail"] = "PMS routing compare queued"
                        });
                    }

                    if (req.Guid.Enabled)
                    {
                        _multiGuidRows.Add(new Dictionary<string, object>
                        {
                            ["File"] = safe,
                            ["Status"] = "CHECK",
                            ["Detail"] = "GUID audit queued",
                            ["IncludeFamily"] = req.Guid.IncludeFamily,
                            ["IncludeAnnotation"] = req.Guid.IncludeAnnotation
                        });
                    }

                    AppendRunItem(safe, "success", string.Empty, "DONE", started);
                }
                catch (Exception ex)
                {
                    AppendRunItem(safe, "failed", ex.Message, "RUN", started);
                    AppendErrorRow(_multiConnectorRows, safe, "파일 처리 실패: " + ex.Message);
                    SendToWeb("host:warn", new { message = $"파일 처리 실패: {safe} - {ex.Message}" });
                }
            }

            var summary = BuildMultiSummaryPayload();
            SendToWeb("hub:multi-done", new { summary });
            SendToWeb("multi:review-summary", summary);
        }

        private void AppendRunItem(string file, string status, string reason, string phase, DateTime started)
        {
            _multiRunItems.Add(new Dictionary<string, object>
            {
                ["File"] = file ?? string.Empty,
                ["Status"] = status ?? string.Empty,
                ["Reason"] = reason ?? string.Empty,
                ["Phase"] = phase ?? string.Empty,
                ["ElapsedMs"] = (long)Math.Max(0, (DateTime.Now - started).TotalMilliseconds)
            });
        }

        private static void AppendErrorRow(List<Dictionary<string, object>> rows, string file, string message)
        {
            rows.Add(new Dictionary<string, object>
            {
                ["File"] = file ?? string.Empty,
                ["Status"] = "ERROR",
                ["Detail"] = message ?? string.Empty
            });
        }

        private object BuildMultiSummaryPayload()
        {
            return new
            {
                connector = new { total = _multiConnectorRows.Count },
                pmsClass = new { total = _multiPmsClassRows.Count },
                pmsSize = new { total = _multiPmsSizeRows.Count },
                pmsRouting = new { total = _multiPmsRoutingRows.Count },
                guid = new { total = _multiGuidRows.Count },
                run = new { total = _multiRunItems.Count }
            };
        }

        private void HandleMultiExport(UIApplication app, object payload)
        {
            try
            {
                var key = GetString(payload, "key", "connector").ToLowerInvariant();
                var rows = key switch
                {
                    "connector" => _multiConnectorRows,
                    "pms" => _multiPmsSizeRows,
                    "pms-class" => _multiPmsClassRows,
                    "pms-routing" => _multiPmsRoutingRows,
                    "segmentpms" => _multiPmsSizeRows,
                    "guid" => _multiGuidRows,
                    "run" => _multiRunItems,
                    _ => null
                };

                if (rows == null)
                {
                    SendToWeb("hub:multi-exported", new { ok = false, message = "알 수 없는 기능 키입니다." });
                    return;
                }
                if (rows.Count == 0)
                {
                    SendToWeb("hub:multi-exported", new { ok = false, message = "내보낼 결과가 없습니다." });
                    return;
                }

                var dt = BuildTableFromRows(rows);
                ExcelCore.EnsureNoDataRow(dt, "오류가 없습니다.");

                var dlg = new Microsoft.Win32.SaveFileDialog { Filter = "Excel (*.xlsx)|*.xlsx", FileName = $"multi_{key}_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx" };
                if (dlg.ShowDialog() != true)
                {
                    SendToWeb("hub:multi-exported", new { ok = false, message = "엑셀 저장이 취소되었습니다." });
                    return;
                }

                var doAutoFit = ParseExcelMode(payload);
                ExcelCore.SaveXlsx(dlg.FileName, dt.TableName, dt, doAutoFit, sheetKey: key, progressKey: "hub:multi-progress", exportKind: key);
                TryApplyExportStyles(key, dlg.FileName, doAutoFit, doAutoFit ? "normal" : "fast");
                SendToWeb("hub:multi-exported", new { ok = true, path = dlg.FileName });
            }
            catch (Exception ex)
            {
                SendToWeb("hub:multi-exported", new { ok = false, message = ex.Message });
            }
        }


        private static double GetDouble(object payload, string key, double def)
        {
            try
            {
                var raw = GetProp(payload, key);
                if (raw == null) return def;
                if (raw is double d) return d;
                if (raw is float f) return f;
                if (raw is decimal m) return (double)m;
                if (raw is int i) return i;
                var text = Convert.ToString(raw);
                if (double.TryParse(text, out var v)) return v;
                if (double.TryParse(text, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out v)) return v;
            }
            catch
            {
            }

            return def;
        }

        private static DataTable BuildTableFromRows(List<Dictionary<string, object>> rows)
        {
            var dt = new DataTable("Multi");
            var cols = rows.SelectMany(r => r.Keys).Distinct(StringComparer.Ordinal).ToList();
            foreach (var c in cols) dt.Columns.Add(c);
            foreach (var r in rows)
            {
                var dr = dt.NewRow();
                foreach (var c in cols)
                {
                    dr[c] = r.TryGetValue(c, out var v) ? Convert.ToString(v) : string.Empty;
                }
                dt.Rows.Add(dr);
            }
            return dt;
        }

        private void TryApplyExportStyles(string exportKey, string savedPath, bool doAutoFit = true, string excelMode = "normal")
        {
            if (string.IsNullOrWhiteSpace(exportKey) || string.IsNullOrWhiteSpace(savedPath)) return;
            try
            {
                ExcelExportStyleRegistry.ApplyStylesForKey(exportKey, savedPath, autoFit: doAutoFit, excelMode: excelMode);
            }
            catch
            {
                // style 실패 시 무시
            }
        }
    }
}
