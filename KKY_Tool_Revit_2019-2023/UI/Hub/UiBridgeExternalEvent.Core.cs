using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading;
using Autodesk.Revit.UI;

namespace KKY_Tool_Revit.UI.Hub
{
    public sealed partial class UiBridgeExternalEvent
    {
        internal static HubHostWindow _host;
        private static readonly UiBridgeExternalEvent Self = new UiBridgeExternalEvent();
        private static readonly object Gate = new object();
        private static readonly Queue<Action<UIApplication>> Queue = new Queue<Action<UIApplication>>();

        private static ExternalEvent _extEv;
        private static IExternalEventHandler _handler;

        public static void Initialize(HubHostWindow host)
        {
            _host = host;

            if (_extEv == null)
            {
                _handler = new BridgeHandler(ProcessQueue);
                _extEv = ExternalEvent.Create(_handler);
            }

            BroadcastTopmost();
            SendToWeb("host:connected", new { ok = true });
        }

        public static void Raise(string name, object payload)
        {
            Enqueue(app => Dispatch(app, name, payload));
        }

        private static void Enqueue(Action<UIApplication> work)
        {
            lock (Gate)
            {
                Queue.Enqueue(work);
            }

            _extEv?.Raise();
        }

        private static void ProcessQueue(UIApplication app)
        {
            while (true)
            {
                Action<UIApplication> todo = null;
                lock (Gate)
                {
                    if (Queue.Count > 0)
                    {
                        todo = Queue.Dequeue();
                    }
                }

                if (todo == null)
                {
                    break;
                }

                try
                {
                    todo(app);
                }
                catch (Exception ex)
                {
                    SendToWeb("host:error", new { message = ex.Message });
                }
            }
        }

        private static void Dispatch(UIApplication app, string name, object payload)
        {
            switch (name)
            {
                case "ui:query-topmost":
                    BroadcastTopmost();
                    return;

                case "ui:set-topmost":
                    var turnOn = false;
                    try
                    {
                        var raw = GetProp(payload, "on");
                        if (raw != null)
                        {
                            turnOn = Convert.ToBoolean(raw);
                        }
                    }
                    catch
                    {
                    }

                    try
                    {
                        if (_host != null)
                        {
                            _host.Topmost = turnOn;
                        }
                    }
                    catch
                    {
                    }

                    BroadcastTopmost();
                    return;

                case "ui:toggle-topmost":
                    try
                    {
                        if (_host != null)
                        {
                            _host.Topmost = !_host.Topmost;
                        }
                    }
                    catch
                    {
                    }

                    BroadcastTopmost();
                    return;
            }

            var map = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                ["dup:run"] = "HandleDupRun",
                ["duplicate:export"] = "HandleDuplicateExport",
                ["duplicate:delete"] = "HandleDuplicateDelete",
                ["duplicate:restore"] = "HandleDuplicateRestore",
                ["duplicate:select"] = "HandleDuplicateSelect",
                ["connector:run"] = "HandleConnectorRun",
                ["connector:save-excel"] = "HandleConnectorSaveExcel",
                ["export:browse-folder"] = "HandleExportBrowse",
                ["export:add-rvt-files"] = "HandleExportAddRvtFiles",
                ["export:preview"] = "HandleExportPreview",
                ["export:save-excel"] = "HandleExportSaveExcel",
                ["paramprop:run"] = "HandleSharedParamRun",
                ["sharedparam:run"] = "HandleSharedParamRun",
                ["sharedparam:list"] = "HandleSharedParamList",
                ["sharedparam:status"] = "HandleSharedParamStatus",
                ["sharedparam:export-excel"] = "HandleSharedParamExport",
                ["sharedparambatch:init"] = "HandleSharedParamBatchInit",
                ["sharedparambatch:browse-rvts"] = "HandleSharedParamBatchBrowseRvts",
                ["sharedparambatch:browse-folder"] = "HandleSharedParamBatchBrowseFolder",
                ["sharedparambatch:run"] = "HandleSharedParamBatchRun",
                ["sharedparambatch:export-excel"] = "HandleSharedParamBatchExportExcel",
                ["sharedparambatch:open-folder"] = "HandleSharedParamBatchOpenFolder",
                ["excel:open"] = "HandleExcelOpen",
                ["segmentpms:rvt-pick-files"] = "HandleSegmentPmsRvtPickFiles",
                ["segmentpms:rvt-pick-folder"] = "HandleSegmentPmsRvtPickFolder",
                ["segmentpms:extract"] = "HandleSegmentPmsExtractStart",
                ["segmentpms:load-extract"] = "HandleSegmentPmsLoadExtract",
                ["segmentpms:save-extract"] = "HandleSegmentPmsSaveExtract",
                ["segmentpms:register-pms"] = "HandleSegmentPmsRegisterPms",
                ["segmentpms:pms-template"] = "HandleSegmentPmsExportTemplate",
                ["segmentpms:prepare-mapping"] = "HandleSegmentPmsPrepareMapping",
                ["segmentpms:run"] = "HandleSegmentPmsRun",
                ["segmentpms:save-result"] = "HandleSegmentPmsSaveResult",
                ["guid:add-files"] = "HandleGuidAddFiles",
                ["guid:run"] = "HandleGuidRun",
                ["guid:export"] = "HandleGuidExport",
                ["guid:request-family-detail"] = "HandleGuidRequestFamilyDetail",
                ["familylink:init"] = "HandleFamilyLinkInit",
                ["familylink:pick-rvts"] = "HandleFamilyLinkPickRvts",
                ["familylink:run"] = "HandleFamilyLinkRun",
                ["familylink:export"] = "HandleFamilyLinkExport",
                ["hub:pick-rvt"] = "HandleMultiPickRvt",
                ["hub:multi-run"] = "HandleMultiRun",
                ["hub:multi-export"] = "HandleMultiExport",
                ["hub:multi-clear"] = "HandleMultiClear",
                ["commonoptions:get"] = "HandleCommonOptionsGet",
                ["commonoptions:save"] = "HandleCommonOptionsSave"
            };

            if (!map.TryGetValue(name, out var methodName))
            {
                SendToWeb("host:warn", new { message = $"알 수 없는 이벤트 '{name}'" });
                return;
            }

            var t = typeof(UiBridgeExternalEvent);
            const BindingFlags flags = BindingFlags.Instance | BindingFlags.NonPublic | BindingFlags.Public;
            var m = t.GetMethod(methodName, flags);

            if (m == null)
            {
                SendToWeb("host:warn", new { message = $"핸들러 '{methodName}' 가 구현되어 있지 않습니다." });
                return;
            }

            var ps = m.GetParameters();
            object[] args;
            switch (ps.Length)
            {
                case 2:
                    args = new object[] { app, payload };
                    break;
                case 1:
                    args = ps[0].ParameterType == typeof(UIApplication)
                        ? new object[] { app }
                        : new object[] { payload };
                    break;
                default:
                    args = Array.Empty<object>();
                    break;
            }

            try
            {
                m.Invoke(Self, args);
            }
            catch (TargetInvocationException ex)
            {
                var msg = ex.InnerException?.Message ?? ex.Message;
                SendToWeb("host:error", new { message = $"핸들러 실행 오류({methodName}): {msg}" });
            }
            catch (Exception ex)
            {
                SendToWeb("host:error", new { message = $"핸들러 실행 오류({methodName}): {ex.Message}" });
            }
        }

        internal static void SendToWeb(string channel, object payload)
        {
            try
            {
                _host?.SendToWeb(channel, payload);
            }
            catch
            {
            }
        }

        private static void BroadcastTopmost()
        {
            try
            {
                var onTop = _host != null && _host.Topmost;
                SendToWeb("host:topmost", new { on = onTop });
            }
            catch
            {
            }
        }

        internal static bool ParseExcelMode(object payload)
        {
            var mode = "normal";

            try
            {
                if (payload != null)
                {
                    var modeProp = GetProp(payload, "excelMode");
                    if (modeProp != null)
                    {
                        var raw = Convert.ToString(modeProp);
                        if (!string.IsNullOrWhiteSpace(raw))
                        {
                            mode = raw;
                        }
                    }
                }
            }
            catch
            {
            }

            return !string.Equals(mode, "normal", StringComparison.OrdinalIgnoreCase);
        }

        private static string GetPropString(object payload, string key)
        {
            var v = GetProp(payload, key);
            return v == null ? null : Convert.ToString(v);
        }

        internal static object GetProp(object payload, string key)
        {
            if (payload == null || string.IsNullOrEmpty(key))
            {
                return null;
            }

            if (payload is IDictionary<string, object> dict)
            {
                return dict.TryGetValue(key, out var v) ? v : null;
            }

            if (payload is IDictionary legacyDict)
            {
                return legacyDict.Contains(key) ? legacyDict[key] : null;
            }

            try
            {
                var t = payload.GetType();
                var p = t.GetProperty(key, BindingFlags.Instance | BindingFlags.Public | BindingFlags.IgnoreCase);
                return p?.GetValue(payload, null);
            }
            catch
            {
                return null;
            }
        }

        internal static string GetString(object payload, string key, string fallback = "")
        {
            try
            {
                var v = GetProp(payload, key);
                if (v == null)
                {
                    return fallback;
                }

                var s = Convert.ToString(v);
                return string.IsNullOrWhiteSpace(s) ? fallback : s;
            }
            catch
            {
                return fallback;
            }
        }

        internal static int GetInt(object payload, string key, int fallback = 0)
        {
            try
            {
                var v = GetProp(payload, key);
                if (v == null)
                {
                    return fallback;
                }

                if (v is int i)
                {
                    return i;
                }

                if (int.TryParse(Convert.ToString(v), NumberStyles.Integer, CultureInfo.InvariantCulture, out var parsed))
                {
                    return parsed;
                }
            }
            catch
            {
            }

            return fallback;
        }

        internal static bool GetBool(object payload, string key, bool fallback = false)
        {
            try
            {
                var v = GetProp(payload, key);
                if (v == null)
                {
                    return fallback;
                }

                if (v is bool b)
                {
                    return b;
                }

                if (bool.TryParse(Convert.ToString(v), out var parsed))
                {
                    return parsed;
                }

                var s = Convert.ToString(v);
                if (string.Equals(s, "1", StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }

                if (string.Equals(s, "0", StringComparison.OrdinalIgnoreCase))
                {
                    return false;
                }
            }
            catch
            {
            }

            return fallback;
        }

        internal static List<string> GetStringList(object payload, string key)
        {
            var res = new List<string>();

            try
            {
                var v = GetProp(payload, key);
                if (v == null)
                {
                    return res;
                }

                if (v is string s)
                {
                    if (!string.IsNullOrWhiteSpace(s))
                    {
                        res.Add(s);
                    }

                    return res;
                }

                if (v is IEnumerable en)
                {
                    foreach (var o in en)
                    {
                        var item = Convert.ToString(o);
                        if (!string.IsNullOrWhiteSpace(item))
                        {
                            res.Add(item);
                        }
                    }
                }
            }
            catch
            {
            }

            return res;
        }

        private sealed class BridgeHandler : IExternalEventHandler
        {
            private readonly Action<UIApplication> _run;

            public BridgeHandler(Action<UIApplication> run)
            {
                _run = run;
            }

            public void Execute(UIApplication app)
            {
                _run?.Invoke(app);
            }

            public string GetName()
            {
                return "KKY.UiBridgeExternalEvent";
            }
        }

        private void HandleExcelOpen(object payload)
        {
            var path = GetPropString(payload, "path");
            if (string.IsNullOrWhiteSpace(path))
            {
                return;
            }

            try
            {
                var p = path.Trim();
                if (!File.Exists(p))
                {
                    SendToWeb("excel:opened", new { ok = false, message = "파일을 찾을 수 없습니다.", path = p });
                    return;
                }

                Process.Start(new ProcessStartInfo
                {
                    FileName = p,
                    UseShellExecute = true
                });

                SendToWeb("excel:opened", new { ok = true, path = p });
            }
            catch (Exception ex)
            {
                SendToWeb("excel:opened", new { ok = false, message = ex.Message, path });
            }
        }

        internal static void ReportProgress(
            string channel,
            string phase,
            string message,
            int current,
            int total,
            double percent,
            bool force = false,
            int minIntervalMs = 33)
        {
            if (!force && minIntervalMs > 0)
            {
                var now = Environment.TickCount;
                if (Math.Abs(now - _lastProgressTick) < minIntervalMs)
                {
                    return;
                }

                _lastProgressTick = now;
            }

            SendToWeb(channel, new
            {
                phase,
                message,
                current,
                total,
                percent = Math.Max(0.0, Math.Min(100.0, percent))
            });
        }

        private static int _lastProgressTick = Environment.TickCount;

        internal static bool IsCancellationRequested(CancellationTokenSource cts)
        {
            try
            {
                return cts != null && cts.IsCancellationRequested;
            }
            catch
            {
                return false;
            }
        }

        internal static DataTable EmptyTable(params string[] cols)
        {
            var dt = new DataTable();
            if (cols != null)
            {
                foreach (var c in cols)
                {
                    dt.Columns.Add(c);
                }
            }

            return dt;
        }

        internal static object ShapeDataTable(DataTable dt)
        {
            var columns = new List<string>();
            var rows = new List<object[]>();

            if (dt != null)
            {
                columns.AddRange(dt.Columns.Cast<DataColumn>().Select(c => c.ColumnName));
                foreach (DataRow r in dt.Rows)
                {
                    rows.Add(columns.Select(c => r[c]).ToArray());
                }
            }

            return new { columns, rows };
        }
    }
}
