using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Windows.Forms;
using Autodesk.Revit.UI;
using KKY_Tool_Revit.Infrastructure;
using KKY_Tool_Revit.Services;

namespace KKY_Tool_Revit.UI.Hub
{
    public sealed class UiBridgeExternalEvent : IExternalEventHandler
    {
        private readonly Queue<(string EventName, IDictionary<string, object> Payload)> _queue = new Queue<(string, IDictionary<string, object>)>();
        private readonly Dictionary<string, Action<UIApplication, IDictionary<string, object>>> _handlers;

        private DataTable _guidProject;
        private DataTable _guidFamilyDetail;
        private DataTable _guidFamilyIndex;
        private string _guidRunId = string.Empty;

        public Action<string, object> HostSender { get; set; }

        public UiBridgeExternalEvent()
        {
            _handlers = new Dictionary<string, Action<UIApplication, IDictionary<string, object>>>(StringComparer.OrdinalIgnoreCase)
            {
                ["guid:add-files"] = HandleGuidAddFiles,
                ["guid:run"] = HandleGuidRun,
                ["guid:request-family-detail"] = HandleGuidRequestFamilyDetail,
                ["guid:export"] = HandleGuidExport,
                ["sharedparam:status"] = HandleSharedParamStatus
            };
        }

        public void Enqueue(string eventName, IDictionary<string, object> payload)
        {
            _queue.Enqueue((eventName, payload ?? new Dictionary<string, object>()));
        }

        public void Execute(UIApplication app)
        {
            while (_queue.Count > 0)
            {
                var item = _queue.Dequeue();
                if (_handlers.TryGetValue(item.EventName, out var handler))
                {
                    handler(app, item.Payload);
                }
            }
        }

        public string GetName() => nameof(UiBridgeExternalEvent);

        private void HandleGuidAddFiles(UIApplication app, IDictionary<string, object> payload)
        {
            var pick = payload.TryGetValue("pick", out var p) ? Convert.ToString(p) : string.Empty;
            if (string.Equals(pick, "folder", StringComparison.OrdinalIgnoreCase))
            {
                using var fbd = new FolderBrowserDialog { Description = "RVT 폴더 선택" };
                if (fbd.ShowDialog() != DialogResult.OK) return;
                var paths = System.IO.Directory.GetFiles(fbd.SelectedPath, "*.rvt", System.IO.SearchOption.TopDirectoryOnly);
                SendToWeb("guid:files", new { paths });
                return;
            }

            using var ofd = new OpenFileDialog { Filter = "Revit Project (*.rvt)|*.rvt", Multiselect = true };
            if (ofd.ShowDialog() != DialogResult.OK) return;
            SendToWeb("guid:files", new { paths = ofd.FileNames });
        }

        private void HandleGuidRun(UIApplication app, IDictionary<string, object> payload)
        {
            try
            {
                var mode = payload.TryGetValue("mode", out var modeObj) ? Convert.ToInt32(modeObj) : 1;
                var includeFamily = payload.TryGetValue("includeFamily", out var includeObj) && Convert.ToBoolean(includeObj);
                var includeAnnotation = payload.TryGetValue("includeAnnotation", out var annotationObj) && Convert.ToBoolean(annotationObj);
                var paths = payload.TryGetValue("rvtPaths", out var pathObj) && pathObj is IEnumerable<object> raw
                    ? raw.Select(x => Convert.ToString(x)).Where(x => !string.IsNullOrWhiteSpace(x)).ToArray()
                    : Array.Empty<string>();

                var result = GuidAuditService.Run(
                    app,
                    mode,
                    paths,
                    (pct, text) => SendToWeb("guid:progress", new { pct, text }),
                    includeAnnotation);

                _guidProject = result.Project;
                _guidFamilyDetail = result.FamilyDetail;
                _guidFamilyIndex = result.FamilyIndex;
                _guidRunId = result.RunId;

                SendToWeb("guid:done", new
                {
                    runId = _guidRunId,
                    project = ShapeTable(_guidProject),
                    familyIndex = ShapeTable(_guidFamilyIndex),
                    includeFamily
                });
            }
            catch (Exception ex)
            {
                SendToWeb("guid:error", new { message = ex.Message });
            }
        }

        private void HandleGuidRequestFamilyDetail(UIApplication app, IDictionary<string, object> payload)
        {
            var runId = payload.TryGetValue("runId", out var runObj) ? Convert.ToString(runObj) : string.Empty;
            var familyName = payload.TryGetValue("familyName", out var famObj) ? Convert.ToString(famObj) : string.Empty;
            var rvtPath = payload.TryGetValue("rvtPath", out var pathObj) ? Convert.ToString(pathObj) : string.Empty;

            if (!string.Equals(runId, _guidRunId, StringComparison.OrdinalIgnoreCase))
            {
                SendToWeb("guid:error", new { message = "이전 실행 결과 요청(runId mismatch)", runId });
                return;
            }

            var filtered = _guidFamilyDetail?.AsEnumerable()
                .Where(r => string.Equals(Convert.ToString(r["FamilyName"]), familyName, StringComparison.OrdinalIgnoreCase)
                         && string.Equals(Convert.ToString(r["RvtPath"]), rvtPath, StringComparison.OrdinalIgnoreCase));

            var table = _guidFamilyDetail?.Clone() ?? GuidAuditService.CreateFamilyDetailTable();
            if (filtered != null)
            {
                foreach (var row in filtered)
                {
                    table.ImportRow(row);
                }
            }

            var shaped = ShapeTable(table, "RvtPath");
            SendToWeb("guid:family-detail", new { runId = _guidRunId, rvtPath, familyName, columns = shaped.columns, rows = shaped.rows });
        }

        private void HandleGuidExport(UIApplication app, IDictionary<string, object> payload)
        {
            SendToWeb("guid:warn", new { message = "C# 변환본에서는 GUID Export가 아직 연결되지 않았습니다." });
        }

        private void HandleSharedParamStatus(UIApplication app, IDictionary<string, object> payload)
        {
            var status = SharedParamReader.ReadStatus(app.Application);
            SendToWeb("sharedparam:status", new
            {
                path = status.Path,
                existsOnDisk = status.ExistsOnDisk,
                canOpen = status.CanOpen,
                isSet = status.IsSet,
                status = status.Status,
                warning = status.Warning,
                errorMessage = status.ErrorMessage
            });
        }

        private (List<string> columns, List<object[]> rows) ShapeTable(DataTable table, params string[] skip)
        {
            var skipSet = new HashSet<string>(skip ?? Array.Empty<string>(), StringComparer.OrdinalIgnoreCase);
            var cols = table?.Columns.Cast<DataColumn>().Select(c => c.ColumnName).Where(c => !skipSet.Contains(c)).ToList() ?? new List<string>();
            var rows = new List<object[]>();
            if (table != null)
            {
                foreach (DataRow row in table.Rows)
                {
                    rows.Add(cols.Select(c => row[c]).ToArray());
                }
            }

            return (cols, rows);
        }

        private void SendToWeb(string evt, object payload) => HostSender?.Invoke(evt, payload);
    }
}
