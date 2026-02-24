using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using Autodesk.Revit.UI;
using KKY_Tool_Revit.Services;

namespace KKY_Tool_Revit.UI.Hub
{
    public sealed partial class UiBridgeExternalEvent
    {
        private readonly List<Dictionary<string, object>> _multiConnectorRows = new List<Dictionary<string, object>>();
        private readonly List<Dictionary<string, object>> _multiSegmentRows = new List<Dictionary<string, object>>();
        private readonly List<Dictionary<string, object>> _multiGuidRows = new List<Dictionary<string, object>>();

        private void HandleMultiPickRvt(UIApplication app, object payload)
        {
            var files = new Microsoft.Win32.OpenFileDialog { Filter = "RVT (*.rvt)|*.rvt", Multiselect = true };
            var selected = files.ShowDialog() == true ? files.FileNames.ToList() : new List<string>();
            SendToWeb("hub:rvt-picked", new { paths = selected });
        }

        private void HandleMultiClear(UIApplication app, object payload)
        {
            _multiConnectorRows.Clear();
            _multiSegmentRows.Clear();
            _multiGuidRows.Clear();
            SendToWeb("multi:review-summary", BuildMultiSummaryPayload());
        }

        private void HandleMultiRun(UIApplication app, object payload)
        {
            var files = GetStringList(payload, "files");
            if (files.Count == 0)
            {
                SendToWeb("hub:multi-error", new { message = "검토할 RVT 파일이 없습니다." });
                return;
            }

            var runConnector = GetBool(payload, "runConnector", true);
            var runSegment = GetBool(payload, "runSegmentPms", false);
            var runGuid = GetBool(payload, "runGuid", false);
            if (!runConnector && !runSegment && !runGuid)
            {
                SendToWeb("hub:multi-error", new { message = "선택된 기능이 없습니다." });
                return;
            }

            _multiConnectorRows.Clear();
            _multiSegmentRows.Clear();
            _multiGuidRows.Clear();

            for (var i = 0; i < files.Count; i++)
            {
                var p = files[i];
                var safe = Path.GetFileName(p);
                SendToWeb("hub:multi-progress", new { current = i + 1, total = files.Count, message = $"처리 중: {safe}" });

                if (runConnector)
                {
                    _multiConnectorRows.Add(new Dictionary<string, object>
                    {
                        ["File"] = safe,
                        ["Status"] = "CHECK",
                        ["Detail"] = "Connector check queued"
                    });
                }

                if (runSegment)
                {
                    _multiSegmentRows.Add(new Dictionary<string, object>
                    {
                        ["File"] = safe,
                        ["Status"] = "CHECK",
                        ["Detail"] = "Segment PMS compare queued"
                    });
                }

                if (runGuid)
                {
                    _multiGuidRows.Add(new Dictionary<string, object>
                    {
                        ["File"] = safe,
                        ["Status"] = "CHECK",
                        ["Detail"] = "GUID audit queued"
                    });
                }
            }

            var summary = BuildMultiSummaryPayload();
            SendToWeb("hub:multi-done", new { summary });
            SendToWeb("multi:review-summary", summary);
        }

        private object BuildMultiSummaryPayload()
        {
            return new
            {
                connector = new { total = _multiConnectorRows.Count },
                segmentpms = new { total = _multiSegmentRows.Count },
                guid = new { total = _multiGuidRows.Count }
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
                    "segmentpms" => _multiSegmentRows,
                    "guid" => _multiGuidRows,
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

                var dt = new DataTable("Multi");
                var cols = rows.SelectMany(r => r.Keys).Distinct(StringComparer.Ordinal).ToList();
                foreach (var c in cols) dt.Columns.Add(c);
                foreach (var r in rows)
                {
                    var dr = dt.NewRow();
                    foreach (var c in cols) dr[c] = r.TryGetValue(c, out var v) ? Convert.ToString(v) : string.Empty;
                    dt.Rows.Add(dr);
                }

                var dlg = new Microsoft.Win32.SaveFileDialog { Filter = "Excel (*.xlsx)|*.xlsx", FileName = $"multi_{key}_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx" };
                if (dlg.ShowDialog() != true)
                {
                    SendToWeb("hub:multi-exported", new { ok = false, message = "엑셀 저장이 취소되었습니다." });
                    return;
                }

                Infrastructure.ExcelCore.SaveXlsx(dlg.FileName, dt.TableName, dt, ParseExcelMode(payload), sheetKey: dt.TableName, progressKey: "hub:multi-progress", exportKind: key);
                SendToWeb("hub:multi-exported", new { ok = true, path = dlg.FileName });
            }
            catch (Exception ex)
            {
                SendToWeb("hub:multi-exported", new { ok = false, message = ex.Message });
            }
        }
    }
}
