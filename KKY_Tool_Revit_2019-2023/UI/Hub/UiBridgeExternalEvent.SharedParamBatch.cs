using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Web.Script.Serialization;
using Autodesk.Revit.UI;
using KKY_Tool_Revit.Services;

namespace KKY_Tool_Revit.UI.Hub
{
    public sealed partial class UiBridgeExternalEvent
    {
        private static SharedParamBatchService.RunResult _lastSharedParamBatchResult;
        private static string _lastSharedParamBatchPayloadJson;

        private void HandleSharedParamBatchInit(UIApplication app, object payload)
        {
            try
            {
                var res = SharedParamBatchService.Init(app);
                SendToWeb("sharedparambatch:init", res);
            }
            catch (Exception ex)
            {
                SendToWeb("sharedparambatch:init", new { ok = false, message = ex.Message });
                SendToWeb("revit:error", new { message = ex.Message });
            }
        }

        private void HandleSharedParamBatchBrowseRvts(UIApplication app, object payload)
        {
            try
            {
                var res = SharedParamBatchService.BrowseRvts();
                SendToWeb("sharedparambatch:rvts-picked", res);
            }
            catch (Exception ex)
            {
                SendToWeb("sharedparambatch:rvts-picked", new { ok = false, message = ex.Message });
                SendToWeb("revit:error", new { message = ex.Message });
            }
        }

        private void HandleSharedParamBatchBrowseFolder(UIApplication app, object payload)
        {
            try
            {
                var res = SharedParamBatchService.BrowseRvtFolder();
                SendToWeb("sharedparambatch:rvts-picked", res);
            }
            catch (Exception ex)
            {
                SendToWeb("sharedparambatch:rvts-picked", new { ok = false, message = ex.Message });
                SendToWeb("revit:error", new { message = ex.Message });
            }
        }

        private void HandleSharedParamBatchRun(UIApplication app, object payload)
        {
            try
            {
                var serializer = new JavaScriptSerializer();
                var payloadJson = serializer.Serialize(payload);
                _lastSharedParamBatchPayloadJson = payloadJson;

                IProgress<object> progress = new Progress<object>(p => { SendToWeb("sharedparambatch:progress", p); });

                var res = SharedParamBatchService.Run(app, payloadJson, progress) as SharedParamBatchService.RunResult;
                if (res != null)
                {
                    res.Logs = FilterSharedParamBatchIssueLogs(res.Logs);
                }

                _lastSharedParamBatchResult = res;

                if (res == null)
                {
                    SendToWeb("sharedparambatch:done", new { ok = false, message = "결과를 확인할 수 없습니다." });
                    return;
                }

                var responsePayload = new
                {
                    ok = res.Ok,
                    message = res.Message,
                    summary = res.Summary ?? new SharedParamBatchService.RunSummary(),
                    logs = (res.Logs ?? new List<SharedParamBatchService.LogEntry>()).Select(l => new
                    {
                        level = l.Level,
                        file = l.File,
                        msg = l.Message
                    }).ToList(),
                    logTextPath = res.LogTextPath
                };

                SendToWeb("sharedparambatch:done", responsePayload);

                if (!res.Ok)
                {
                    SendToWeb("revit:error", new { message = "Shared Parameter Batch 실패: " + res.Message });
                }
            }
            catch (Exception ex)
            {
                SendToWeb("sharedparambatch:done", new { ok = false, message = ex.Message });
                SendToWeb("revit:error", new { message = "Shared Parameter Batch 실패: " + ex.Message });
            }
        }

        private static List<SharedParamBatchService.LogEntry> FilterSharedParamBatchIssueLogs(List<SharedParamBatchService.LogEntry> logs)
        {
            var source = logs ?? new List<SharedParamBatchService.LogEntry>();
            var dt = new DataTable("SharedParamBatchLogs");
            dt.Columns.Add("__idx", typeof(int));
            dt.Columns.Add("성공여부");
            dt.Columns.Add("메시지");

            var index = 0;
            foreach (var l in source)
            {
                var row = dt.NewRow();
                row["__idx"] = index;
                row["성공여부"] = l == null ? "" : (l.Level ?? "");
                row["메시지"] = l == null ? "" : (l.Message ?? "");
                dt.Rows.Add(row);
                index += 1;
            }

            var filtered = FilterIssueRowsCopy("sharedparambatch", dt);
            var result = new List<SharedParamBatchService.LogEntry>();
            if (filtered == null) return result;

            foreach (DataRow row in filtered.Rows)
            {
                int sourceIdx;
                try
                {
                    sourceIdx = Convert.ToInt32(row["__idx"]);
                }
                catch
                {
                    sourceIdx = -1;
                }

                if (sourceIdx >= 0 && sourceIdx < source.Count)
                {
                    result.Add(source[sourceIdx]);
                }
            }

            return result;
        }

        private void HandleSharedParamBatchExportExcel(UIApplication app, object payload)
        {
            try
            {
                if (_lastSharedParamBatchResult == null)
                {
                    SendToWeb("sharedparambatch:exported", new { ok = false, message = "최근 실행 결과가 없습니다." });
                    return;
                }

                string mode = null;
                try
                {
                    var prop = GetPropLocal(payload, "excelMode");
                    if (prop != null) mode = Convert.ToString(prop);
                }
                catch
                {
                    // ignore
                }

                var logsPayload = _lastSharedParamBatchResult.Logs.Select(l => new
                {
                    level = l.Level,
                    file = l.File,
                    msg = l.Message
                }).ToList();

                var serializer = new JavaScriptSerializer();
                Dictionary<string, object> runPayload = null;
                if (!string.IsNullOrWhiteSpace(_lastSharedParamBatchPayloadJson))
                {
                    try
                    {
                        runPayload = serializer.Deserialize<Dictionary<string, object>>(_lastSharedParamBatchPayloadJson);
                    }
                    catch
                    {
                        runPayload = null;
                    }
                }

                object rvtPaths = null;
                object parameters = null;
                if (runPayload != null)
                {
                    runPayload.TryGetValue("rvtPaths", out rvtPaths);
                    runPayload.TryGetValue("parameters", out parameters);
                }

                var payloadJson = serializer.Serialize(new
                {
                    logs = logsPayload,
                    excelMode = mode,
                    rvtPaths,
                    parameters
                });

                var res = SharedParamBatchService.ExportExcel(payloadJson);
                SendToWeb("sharedparambatch:exported", res);
            }
            catch (Exception ex)
            {
                SendToWeb("sharedparambatch:exported", new { ok = false, message = ex.Message });
                SendToWeb("revit:error", new { message = "엑셀 내보내기 실패: " + ex.Message });
            }
        }

        private void HandleSharedParamBatchOpenFolder(UIApplication app, object payload)
        {
            var inputPath = GetPropLocal(payload, "path") as string;
            if (string.IsNullOrWhiteSpace(inputPath))
            {
                SendToWeb("sharedparambatch:open-folder", new { ok = false, message = "경로가 비어 있습니다." });
                return;
            }

            var targetPath = inputPath;
            try
            {
                if (File.Exists(inputPath))
                {
                    targetPath = Path.GetDirectoryName(inputPath);
                }
            }
            catch
            {
                // ignore
            }

            if (string.IsNullOrWhiteSpace(targetPath))
            {
                SendToWeb("sharedparambatch:open-folder", new { ok = false, message = "폴더 경로를 확인할 수 없습니다." });
                return;
            }

            try
            {
                var targetPathText = targetPath.ToString();
                var psi = new ProcessStartInfo("explorer.exe", "\"" + targetPathText + "\"")
                {
                    UseShellExecute = true
                };
                Process.Start(psi);
                SendToWeb("sharedparambatch:open-folder", new { ok = true, path = targetPathText });
            }
            catch (Exception ex)
            {
                SendToWeb("sharedparambatch:open-folder", new { ok = false, message = ex.Message });
            }
        }

        private static object GetPropLocal(object obj, string prop)
        {
            if (obj == null) return null;

            var d = obj as IDictionary<string, object>;
            if (d != null)
            {
                object v;
                if (d.TryGetValue(prop, out v)) return v;
                return null;
            }

            var t = obj.GetType();
            var pInfo = t.GetProperty(prop, System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.IgnoreCase);
            if (pInfo == null) return null;
            return pInfo.GetValue(obj, null);
        }
    }
}
