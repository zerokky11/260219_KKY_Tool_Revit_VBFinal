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
    public sealed partial class UiBridgeExternalEvent
    {
        private DataTable _guidProject;
        private DataTable _guidFamilyDetail;
        private DataTable _guidFamilyIndex;
        private string _guidRunId = string.Empty;

        private void HandleGuidAddFiles(UIApplication app, object payload)
        {
            var pick = GetString(payload, "pick", string.Empty);
            if (string.Equals(pick, "folder", StringComparison.OrdinalIgnoreCase))
            {
                using (var fbd = new FolderBrowserDialog { Description = "RVT 폴더 선택" })
                {
                    if (fbd.ShowDialog() != DialogResult.OK)
                    {
                        return;
                    }

                    var paths = System.IO.Directory.GetFiles(fbd.SelectedPath, "*.rvt", System.IO.SearchOption.TopDirectoryOnly);
                    SendToWeb("guid:files", new { paths });
                    return;
                }
            }

            using (var ofd = new OpenFileDialog { Filter = "Revit Project (*.rvt)|*.rvt", Multiselect = true })
            {
                if (ofd.ShowDialog() != DialogResult.OK)
                {
                    return;
                }

                SendToWeb("guid:files", new { paths = ofd.FileNames });
            }
        }

        private void HandleGuidRun(UIApplication app, object payload)
        {
            try
            {
                var mode = GetInt(payload, "mode", 1);
                var includeFamily = GetBool(payload, "includeFamily", false);
                var includeAnnotation = GetBool(payload, "includeAnnotation", false);
                var paths = GetStringList(payload, "rvtPaths").Where(x => !string.IsNullOrWhiteSpace(x)).ToArray();

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

        private void HandleGuidRequestFamilyDetail(UIApplication app, object payload)
        {
            var runId = GetString(payload, "runId", string.Empty);
            var familyName = GetString(payload, "familyName", string.Empty);
            var rvtPath = GetString(payload, "rvtPath", string.Empty);

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

        private void HandleGuidExport(UIApplication app, object payload)
        {
            SendToWeb("guid:warn", new { message = "C# 변환본에서는 GUID Export가 아직 연결되지 않았습니다." });
        }

        private void HandleSharedParamStatus(UIApplication app, object payload)
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
    }
}
