using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
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
        private int _guidMode = 1;
        private string _guidRunId = string.Empty;
        private DataTable _guidFamilyIndex;
        private bool _guidIncludeFamily;
        private double _guidLastPct = -1.0;
        private string _guidLastText = string.Empty;

        private sealed class TablePayload
        {
            public List<string> columns { get; set; }
            public List<object[]> rows { get; set; }
        }

        private void HandleGuidAddFiles(UIApplication app, object payload)
        {
            var pick = GetString(payload, "pick", string.Empty);
            if (string.Equals(pick, "folder", StringComparison.OrdinalIgnoreCase))
            {
                HandleGuidAddFolder();
                return;
            }

            using (var ofd = new OpenFileDialog())
            {
                ofd.Filter = "Revit Project (*.rvt)|*.rvt";
                ofd.Multiselect = true;
                ofd.Title = "검토할 RVT 파일 선택";
                ofd.RestoreDirectory = true;
                if (ofd.ShowDialog() != DialogResult.OK) return;
                SendToWeb("guid:files", new { paths = ofd.FileNames.Where(x => !string.IsNullOrWhiteSpace(x)).ToList() });
            }
        }

        private void HandleGuidRun(UIApplication app, object payload)
        {
            var mode = GetInt(payload, "mode", 1);
            if (mode != 1 && mode != 2) mode = 1;

            var includeFamily = GetBool(payload, "includeFamily", mode == 2);
            var includeAnnotation = GetBool(payload, "includeAnnotation", false);
            var paths = GetStringList(payload, "rvtPaths").Where(x => !string.IsNullOrWhiteSpace(x)).ToArray();

            try
            {
                var sharedStatus = SharedParameterStatusService.GetStatus(app);
                if (sharedStatus == null || !string.Equals(sharedStatus.Status, "ok", StringComparison.OrdinalIgnoreCase))
                {
                    var msg = string.IsNullOrWhiteSpace(sharedStatus?.WarningMessage) ? "Shared Parameter 파일 상태가 올바르지 않습니다." : sharedStatus.WarningMessage;
                    SendToWeb("sharedparam:status", new
                    {
                        path = sharedStatus?.Path,
                        isSet = sharedStatus?.IsSet,
                        existsOnDisk = sharedStatus?.ExistsOnDisk,
                        canOpen = sharedStatus?.CanOpen,
                        status = sharedStatus?.Status,
                        statusLabel = sharedStatus?.StatusLabel,
                        warning = sharedStatus?.WarningMessage,
                        errorMessage = sharedStatus?.ErrorMessage
                    });
                    SendToWeb("guid:error", new { message = msg });
                    SendToWeb("revit:error", new { message = "GUID 검토 실패: " + msg });
                    return;
                }

                _guidProject = null;
                _guidFamilyDetail = null;
                _guidFamilyIndex = null;
                _guidRunId = string.Empty;
                _guidIncludeFamily = includeFamily;
                _guidMode = mode;
                _guidLastPct = -1.0;
                _guidLastText = string.Empty;

                var result = GuidAuditService.Run(
                    app,
                    mode,
                    paths,
                    ReportGuidProgress,
                    msg =>
                    {
                        if (!string.IsNullOrWhiteSpace(msg))
                        {
                            SendToWeb("guid:warn", new { message = msg });
                        }
                    },
                    includeFamily,
                    includeAnnotation);

                _guidProject = FilterIssueRowsCopy("guid", result.Project);
                _guidFamilyDetail = FilterIssueRowsCopy("guid", result.FamilyDetail);
                _guidFamilyIndex = result.FamilyIndex;
                _guidRunId = result.RunId;
                _guidIncludeFamily = result.IncludeFamily;

                var payloadProject = ShapeGuidTable(_guidProject, new HashSet<string>(StringComparer.OrdinalIgnoreCase) { "RvtPath" });
                var payloadFamily = ShapeGuidTable(result.FamilyIndex, null);

                SendToWeb("guid:done", new
                {
                    mode,
                    runId = _guidRunId,
                    includeFamily = _guidIncludeFamily,
                    includeAnnotation,
                    project = payloadProject,
                    family = payloadFamily
                });
            }
            catch (Exception ex)
            {
                SendToWeb("guid:error", new { message = ex.Message });
            }
            finally
            {
                ReportGuidProgress(0, string.Empty);
            }
        }

        private void HandleGuidExport(UIApplication app, object payload)
        {
            var which = GetString(payload, "which", string.Empty).ToLowerInvariant();
            var excelMode = GetString(payload, "excelMode", "fast");

            try
            {
                var projectTable = EnsureNoRvtPath(GuidAuditService.PrepareExportTable(_guidProject, 1));
                var familyTable = EnsureNoRvtPath(GuidAuditService.PrepareExportTable(_guidFamilyDetail, 2));
                var sheets = new List<KeyValuePair<string, DataTable>>();

                if (which == "family")
                {
                    sheets.Add(new KeyValuePair<string, DataTable>("Family 검토결과", familyTable));
                }
                else if (which == "all")
                {
                    sheets.Add(new KeyValuePair<string, DataTable>("RVT 검토결과", projectTable));
                    sheets.Add(new KeyValuePair<string, DataTable>("Family 검토결과", familyTable));
                }
                else
                {
                    sheets.Add(new KeyValuePair<string, DataTable>("RVT 검토결과", projectTable));
                }

                var doAutoFit = ParseExcelMode(payload);
                var exportMode = string.Equals(excelMode, "fast", StringComparison.OrdinalIgnoreCase)
                    ? "fast"
                    : (doAutoFit ? "normal" : "fast");

                LogAutoFitDecision(doAutoFit, "GuidAuditExport");
                var saved = GuidAuditService.ExportMulti(sheets, exportMode, "guid:progress");
                if (string.IsNullOrWhiteSpace(saved))
                {
                    SendToWeb("guid:error", new { message = "엑셀 내보내기가 취소되었습니다." });
                    return;
                }

                ExcelExportStyleRegistry.ApplyStylesForKey("guid", saved, autoFit: doAutoFit, excelMode: exportMode);
                SendToWeb("guid:exported", new { path = saved, which });
            }
            catch (Exception ex)
            {
                SendToWeb("guid:error", new { message = "엑셀 내보내기 실패: " + ex.Message });
            }
        }

        private void HandleGuidRequestFamilyDetail(UIApplication app, object payload)
        {
            var runId = GetString(payload, "runId", string.Empty);
            var familyName = GetString(payload, "familyName", string.Empty);
            var rvtPath = GetString(payload, "rvtPath", string.Empty);

            if (!_guidIncludeFamily)
            {
                SendToWeb("guid:error", new { message = "Family 검토 결과가 없습니다.", runId });
                return;
            }

            if (string.IsNullOrWhiteSpace(runId) || !string.Equals(runId, _guidRunId, StringComparison.OrdinalIgnoreCase))
            {
                SendToWeb("guid:error", new { message = "이전 실행 결과 요청(runId mismatch)", runId });
                return;
            }

            if (_guidFamilyDetail == null || _guidFamilyDetail.Rows.Count == 0)
            {
                SendToWeb("guid:error", new { message = "가져올 패밀리 상세 결과가 없습니다.", runId });
                return;
            }

            var filtered = FilterFamilyDetail(_guidFamilyDetail, rvtPath, familyName);
            var shaped = ShapeGuidTable(filtered, new HashSet<string>(StringComparer.OrdinalIgnoreCase) { "RvtPath" });
            SendToWeb("guid:family-detail", new { runId = _guidRunId, rvtPath, familyName, columns = shaped.columns, rows = shaped.rows });
        }

        private void ReportGuidProgress(double pct, string text)
        {
            var changed = (pct != _guidLastPct) || !string.Equals(text, _guidLastText, StringComparison.Ordinal);
            if (!changed) return;
            SendToWeb("guid:progress", new { pct, text });
            _guidLastPct = pct;
            _guidLastText = text;
        }

        private TablePayload ShapeGuidTable(DataTable dt, HashSet<string> skipCols)
        {
            if (dt == null)
            {
                return new TablePayload { columns = new List<string>(), rows = new List<object[]>() };
            }

            var cols = new List<string>();
            foreach (DataColumn c in dt.Columns)
            {
                if (skipCols != null && skipCols.Contains(c.ColumnName)) continue;
                cols.Add(c.ColumnName);
            }

            var rows = new List<object[]>();
            foreach (DataRow r in dt.Rows)
            {
                var arr = new object[cols.Count];
                for (var i = 0; i < cols.Count; i++) arr[i] = SafeStrGuid(r[cols[i]]);
                rows.Add(arr);
            }

            return new TablePayload { columns = cols, rows = rows };
        }

        private DataTable CloneWithoutColumn(DataTable dt, string columnName)
        {
            if (dt == null) return null;
            var clone = dt.Clone();
            if (clone.Columns.Contains(columnName)) clone.Columns.Remove(columnName);
            foreach (DataRow r in dt.Rows)
            {
                var nr = clone.NewRow();
                foreach (DataColumn c in clone.Columns)
                {
                    nr[c.ColumnName] = r[c.ColumnName];
                }
                clone.Rows.Add(nr);
            }
            return clone;
        }

        private DataTable EnsureNoRvtPath(DataTable dt)
        {
            if (dt == null) return null;
            return dt.Columns.Contains("RvtPath") ? CloneWithoutColumn(dt, "RvtPath") : dt;
        }

        private static string SafeStrGuid(object o)
        {
            if (o == null || o == DBNull.Value) return string.Empty;
            return Convert.ToString(o);
        }

        private DataTable FilterFamilyDetail(DataTable source, string rvtPath, string familyName)
        {
            if (source == null) return new DataTable();
            var clone = source.Clone();
            var path = rvtPath ?? string.Empty;
            var fam = familyName ?? string.Empty;
            foreach (DataRow r in source.Rows)
            {
                var rp = SafeStrGuid(r["RvtPath"]);
                var fn = SafeStrGuid(r["FamilyName"]);
                if (string.Equals(rp, path, StringComparison.OrdinalIgnoreCase) &&
                    string.Equals(fn, fam, StringComparison.OrdinalIgnoreCase))
                {
                    clone.ImportRow(r);
                }
            }
            return clone;
        }

        private void HandleGuidAddFolder()
        {
            using (var dlg = new FolderBrowserDialog())
            {
                dlg.Description = "RVT 폴더 선택";
                dlg.ShowNewFolderButton = false;
                if (dlg.ShowDialog() != DialogResult.OK) return;

                var root = dlg.SelectedPath;
                const int maxFiles = 2000;
                var files = new List<string>();

                try
                {
                    if (Directory.Exists(root))
                    {
                        var found = Directory.EnumerateFiles(root, "*.rvt", SearchOption.TopDirectoryOnly)
                            .Select(p => new { Path = p, Name = Path.GetFileName(p) })
                            .OrderBy(x => x.Name, StringComparer.OrdinalIgnoreCase)
                            .ToList();

                        foreach (var item in found)
                        {
                            if (string.IsNullOrWhiteSpace(item.Path)) continue;
                            if (files.Count >= maxFiles) break;
                            files.Add(item.Path);
                        }

                        if (found.Count > maxFiles)
                        {
                            SendToWeb("guid:warn", new { message = $"RVT 파일이 {found.Count:#,0}개 있습니다. 상위 {maxFiles:#,0}개만 추가합니다." });
                        }
                    }
                }
                catch (Exception ex)
                {
                    SendToWeb("guid:warn", new { message = $"폴더를 읽는 중 오류가 발생했습니다: {ex.Message}" });
                }

                if (files.Count == 0)
                {
                    SendToWeb("guid:warn", new { message = "선택한 폴더에 RVT 파일이 없습니다." });
                    return;
                }

                var deduped = files.Distinct(StringComparer.OrdinalIgnoreCase).ToList();
                SendToWeb("guid:files", new { paths = deduped });
            }
        }
    }
}
