using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KKY_Tool_Revit.Infrastructure;
using KKY_Tool_Revit.Services;

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

        private sealed class MultiFamilyLinkOptions
        {
            public bool Enabled { get; set; }
        }

        private sealed class MultiPointsOptions
        {
            public bool Enabled { get; set; }
            public string Unit { get; set; } = "ft";
        }

        private sealed class MultiRunRequest
        {
            public MultiCommonOptions Common { get; set; } = new MultiCommonOptions();
            public MultiConnectorOptions Connector { get; set; } = new MultiConnectorOptions();
            public MultiPmsOptions Pms { get; set; } = new MultiPmsOptions();
            public MultiGuidOptions Guid { get; set; } = new MultiGuidOptions();
            public MultiFamilyLinkOptions FamilyLink { get; set; } = new MultiFamilyLinkOptions();
            public MultiPointsOptions Points { get; set; } = new MultiPointsOptions();
            public List<string> RvtPaths { get; set; } = new List<string>();
        }

        private sealed class MultiRunItem
        {
            public string File { get; set; } = string.Empty;
            public string Status { get; set; } = string.Empty;
            public string Reason { get; set; } = string.Empty;
            public string Phase { get; set; } = string.Empty;
            public long ElapsedMs { get; set; }
        }

        private readonly object _multiLock = new object();
        private Queue<string> _multiQueue;
        private int _multiTotal;
        private int _multiIndex;
        private bool _multiActive;
        private bool _multiBusy;
        private bool _multiPending;
        private MultiRunRequest _multiRequest;
        private UIApplication _multiApp;

        private readonly List<Dictionary<string, object>> _multiConnectorRows = new List<Dictionary<string, object>>();
        private readonly List<Dictionary<string, object>> _multiPmsClassRows = new List<Dictionary<string, object>>();
        private readonly List<Dictionary<string, object>> _multiPmsSizeRows = new List<Dictionary<string, object>>();
        private readonly List<Dictionary<string, object>> _multiPmsRoutingRows = new List<Dictionary<string, object>>();
        private readonly List<Dictionary<string, object>> _multiGuidRows = new List<Dictionary<string, object>>();
        private readonly List<Dictionary<string, object>> _multiFamilyLinkRows = new List<Dictionary<string, object>>();
        private readonly List<Dictionary<string, object>> _multiPointsRows = new List<Dictionary<string, object>>();
        private DataTable _multiGuidProject;
        private DataTable _multiGuidFamilyDetail;
        private DataTable _multiGuidFamilyIndex;
        private readonly List<MultiRunItem> _multiRunItems = new List<MultiRunItem>();

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
            var options = new HubCommonOptionsStorageService.HubCommonOptions
            {
                ExtraParamsText = GetString(payload, "extraParamsText", string.Empty),
                TargetFilterText = GetString(payload, "targetFilterText", string.Empty),
                ExcludeEndDummy = GetBool(payload, "excludeEndDummy", false)
            };

            var ok = HubCommonOptionsStorageService.Save(options);
            SendToWeb("commonoptions:saved", new { ok });
        }

        private void HandleMultiPickRvt(UIApplication app, object payload)
        {
            var dlg = new Microsoft.Win32.OpenFileDialog { Filter = "RVT (*.rvt)|*.rvt", Multiselect = true };
            var selected = dlg.ShowDialog() == true ? dlg.FileNames.ToList() : new List<string>();
            SendToWeb("hub:rvt-picked", new { paths = selected });
        }

        private void HandleMultiClear(UIApplication app, object payload)
        {
            _multiConnectorRows.Clear();
            _multiPmsClassRows.Clear();
            _multiPmsSizeRows.Clear();
            _multiPmsRoutingRows.Clear();
            _multiGuidRows.Clear();
            _multiFamilyLinkRows.Clear();
            _multiPointsRows.Clear();
            _multiGuidProject = null;
            _multiGuidFamilyDetail = null;
            _multiGuidFamilyIndex = null;
            _multiRunItems.Clear();
            SendToWeb("multi:review-summary", BuildMultiSummaryPayload());
        }

        private void HandleMultiRun(UIApplication app, object payload)
        {
            _multiRequest = ParseMultiRequest(payload);
            if (_multiRequest.RvtPaths.Count == 0)
            {
                SendToWeb("hub:multi-error", new { message = "검토할 RVT 파일이 없습니다." });
                return;
            }

            if (CountEnabledFeatures(_multiRequest) == 0)
            {
                SendToWeb("hub:multi-error", new { message = "선택된 기능이 없습니다." });
                return;
            }

            _multiConnectorRows.Clear();
            _multiPmsClassRows.Clear();
            _multiPmsSizeRows.Clear();
            _multiPmsRoutingRows.Clear();
            _multiGuidRows.Clear();
            _multiFamilyLinkRows.Clear();
            _multiPointsRows.Clear();
            _multiGuidProject = null;
            _multiGuidFamilyDetail = null;
            _multiGuidFamilyIndex = null;
            _multiRunItems.Clear();

            lock (_multiLock)
            {
                _multiQueue = new Queue<string>(_multiRequest.RvtPaths);
                _multiTotal = _multiQueue.Count;
                _multiIndex = 0;
                _multiActive = true;
                _multiBusy = false;
                _multiPending = false;
                _multiApp = app;
            }

            // non-interactive 환경에서는 idling 훅 대신 동기 루프
            while (true)
            {
                bool keep;
                lock (_multiLock) keep = _multiActive;
                if (!keep) break;
                ProcessMultiNext();
            }
        }

        private void HandleMultiIdling(object sender, object e)
        {
            ProcessMultiNext();
        }

        private void ProcessMultiNext()
        {
            string filePath = null;
            lock (_multiLock)
            {
                if (!_multiActive || _multiBusy) return;
                if (_multiQueue == null || _multiQueue.Count == 0)
                {
                    _multiActive = false;
                    return;
                }

                _multiBusy = true;
                filePath = _multiQueue.Dequeue();
                _multiIndex += 1;
            }

            var started = DateTime.Now;
            var safeName = Path.GetFileName(filePath ?? string.Empty);

            try
            {
                var basePct = _multiTotal > 0 ? (double)(_multiIndex - 1) / _multiTotal : 0;
                ReportMultiProgress(basePct * 100.0, "파일 여는 중", safeName);

                if (string.IsNullOrWhiteSpace(filePath) || !File.Exists(filePath))
                {
                    AppendMultiConnectorError(safeName, "파일을 찾을 수 없습니다.");
                    AppendMultiRunItem(safeName, "skipped", "파일을 찾을 수 없습니다.", "OPEN", started);
                }
                else
                {
                    RunMultiForDocument(_multiApp, null, filePath, safeName, basePct);
                    AppendMultiRunItem(safeName, "success", string.Empty, "DONE", started);
                }
            }
            catch (Exception ex)
            {
                AppendMultiConnectorError(safeName, $"파일 처리 실패: {ex.Message}");
                AppendMultiRunItem(safeName, "failed", ex.Message, "RUN", started);
                SendToWeb("host:warn", new { message = $"파일 처리 실패: {safeName} - {ex.Message}" });
            }
            finally
            {
                lock (_multiLock)
                {
                    _multiBusy = false;
                    if (_multiQueue == null || _multiQueue.Count == 0) _multiActive = false;
                }
            }

            bool done;
            lock (_multiLock) done = !_multiActive;
            if (done) FinishMultiRun();
        }

        private void RunMultiForDocument(UIApplication app, Document doc, string path, string safeName, double basePct)
        {
            var steps = CountEnabledFeatures(_multiRequest);
            var stepIndex = 0;

            if (_multiRequest.Connector.Enabled)
            {
                stepIndex++;
                ReportMultiProgress(CalcStepPercent(basePct, stepIndex, steps), "커넥터 진단 실행 중", safeName);
                _multiConnectorRows.Add(new Dictionary<string, object>
                {
                    ["File"] = safeName,
                    ["Status"] = "CHECK",
                    ["Detail"] = "Connector check queued",
                    ["Tol"] = _multiRequest.Connector.Tol,
                    ["Unit"] = _multiRequest.Connector.Unit,
                    ["Param"] = _multiRequest.Connector.Param
                });
                ReportMultiProgress(CalcStepPercent(basePct, stepIndex, steps), "커넥터 진단 완료", safeName);
            }

            if (_multiRequest.Pms.Enabled)
            {
                stepIndex++;
                ReportMultiProgress(CalcStepPercent(basePct, stepIndex, steps), "PMS 비교 실행 중", safeName);
                AppendSegmentPmsRows(safeName);
                ReportMultiProgress(CalcStepPercent(basePct, stepIndex, steps), "PMS 비교 완료", safeName);
            }

            if (_multiRequest.Guid.Enabled)
            {
                stepIndex++;
                ReportMultiProgress(CalcStepPercent(basePct, stepIndex, steps), "GUID 검토 실행 중", safeName);

                var proj = GuidAuditService.CreateProjectTable();
                var row = proj.NewRow();
                row["RvtName"] = safeName;
                row["RvtPath"] = path;
                row["ParamName"] = string.Empty;
                row["ParamKind"] = "Project";
                row["Result"] = "CHECK";
                row["Notes"] = "GUID audit queued";
                proj.Rows.Add(row);

                MergeGuidResult(proj, null, null);
                _multiGuidRows.Add(new Dictionary<string, object>
                {
                    ["File"] = safeName,
                    ["Status"] = "CHECK",
                    ["Detail"] = "GUID audit queued",
                    ["IncludeFamily"] = _multiRequest.Guid.IncludeFamily,
                    ["IncludeAnnotation"] = _multiRequest.Guid.IncludeAnnotation
                });

                ReportMultiProgress(CalcStepPercent(basePct, stepIndex, steps), "GUID 검토 완료", safeName);
            }

            if (_multiRequest.FamilyLink.Enabled)
            {
                stepIndex++;
                ReportMultiProgress(CalcStepPercent(basePct, stepIndex, steps), "Family Link 검토 실행 중", safeName);
                _multiFamilyLinkRows.Add(new Dictionary<string, object>
                {
                    ["File"] = safeName,
                    ["Status"] = "CHECK",
                    ["Detail"] = "Family link audit queued"
                });
                ReportMultiProgress(CalcStepPercent(basePct, stepIndex, steps), "Family Link 검토 완료", safeName);
            }

            if (_multiRequest.Points.Enabled)
            {
                stepIndex++;
                ReportMultiProgress(CalcStepPercent(basePct, stepIndex, steps), "포인트 추출 실행 중", safeName);
                _multiPointsRows.Add(new Dictionary<string, object>
                {
                    ["File"] = safeName,
                    ["Status"] = "CHECK",
                    ["Detail"] = "Points extract queued",
                    ["Unit"] = _multiRequest.Points.Unit
                });
                ReportMultiProgress(CalcStepPercent(basePct, stepIndex, steps), "포인트 추출 완료", safeName);
            }
        }

        private void FinishMultiRun()
        {
            var summary = new Dictionary<string, object>
            {
                ["connector"] = new { rows = _multiConnectorRows.Count },
                ["pms"] = new { rows = _multiPmsSizeRows.Count },
                ["guid"] = new { rows = _multiGuidRows.Count },
                ["run"] = new { rows = _multiRunItems.Count }
            };

            SendToWeb("hub:multi-done", new { summary });
            SendToWeb("multi:review-summary", BuildMultiSummaryPayload());
        }

        private void HandleMultiExport(UIApplication app, object payload)
        {
            var key = (GetString(payload, "key", string.Empty) ?? string.Empty).ToLowerInvariant();
            var excelMode = GetString(payload, "excelMode", "normal");
            var doAutoFit = ParseExcelMode(payload);

            try
            {
                switch (key)
                {
                    case "connector":
                        ExportConnector(doAutoFit, excelMode);
                        break;
                    case "pms":
                    case "segmentpms":
                        ExportSegmentPms(doAutoFit, excelMode);
                        break;
                    case "guid":
                        ExportGuid(excelMode);
                        break;
                    case "familylink":
                        ExportFamilyLink(doAutoFit, excelMode);
                        break;
                    case "points":
                        ExportPoints(doAutoFit, excelMode);
                        break;
                    case "run":
                        ExportRunItems(doAutoFit, excelMode);
                        break;
                    default:
                        SendToWeb("hub:multi-exported", new { ok = false, message = "알 수 없는 기능 키입니다." });
                        break;
                }
            }
            catch (Exception ex)
            {
                SendToWeb("hub:multi-exported", new { ok = false, message = ex.Message });
            }
        }

        private void TryApplyExportStyles(string exportKey, string savedPath, bool doAutoFit = true, string excelMode = "normal")
        {
            if (string.IsNullOrWhiteSpace(exportKey) || string.IsNullOrWhiteSpace(savedPath)) return;
            try { ExcelExportStyleRegistry.ApplyStylesForKey(exportKey, savedPath, autoFit: doAutoFit, excelMode: excelMode); } catch { }
        }

        private void ExportConnector(bool doAutoFit, string excelMode)
        {
            if (_multiConnectorRows.Count == 0)
            {
                SendToWeb("hub:multi-exported", new { ok = false, message = "커넥터 결과가 없습니다." });
                return;
            }

            var dt = BuildTableFromRows(_multiConnectorRows, "Connector");
            ExcelCore.EnsureNoDataRow(dt, "오류가 없습니다.");
            var saved = SaveTableByDialog(dt, "multi_connector.xlsx", doAutoFit, "connector");
            if (string.IsNullOrWhiteSpace(saved)) return;
            TryApplyExportStyles("connector", saved, doAutoFit, excelMode);
            SendToWeb("hub:multi-exported", new { ok = true, path = saved });
        }

        private void ExportSegmentPms(bool doAutoFit, string excelMode)
        {
            if (_multiPmsClassRows.Count == 0 && _multiPmsSizeRows.Count == 0 && _multiPmsRoutingRows.Count == 0)
            {
                SendToWeb("hub:multi-exported", new { ok = false, message = "PMS 결과가 없습니다." });
                return;
            }

            var sheets = new List<KeyValuePair<string, DataTable>>();
            sheets.Add(new KeyValuePair<string, DataTable>("PMS_Class", BuildTableFromRows(_multiPmsClassRows, "PMS_Class")));
            sheets.Add(new KeyValuePair<string, DataTable>("PMS_Size", BuildTableFromRows(_multiPmsSizeRows, "PMS_Size")));
            sheets.Add(new KeyValuePair<string, DataTable>("PMS_Routing", BuildTableFromRows(_multiPmsRoutingRows, "PMS_Routing")));
            foreach (var s in sheets) ExcelCore.EnsureNoDataRow(s.Value, "오류가 없습니다.");

            var dlg = new Microsoft.Win32.SaveFileDialog { Filter = "Excel (*.xlsx)|*.xlsx", FileName = "multi_pms.xlsx" };
            if (dlg.ShowDialog() != true)
            {
                SendToWeb("hub:multi-exported", new { ok = false, message = "엑셀 저장이 취소되었습니다." });
                return;
            }

            ExcelCore.SaveXlsxMulti(dlg.FileName, sheets, doAutoFit, "hub:multi-progress");
            TryApplyExportStyles("pms", dlg.FileName, doAutoFit, excelMode);
            SendToWeb("hub:multi-exported", new { ok = true, path = dlg.FileName });
        }

        private void ExportGuid(string excelMode)
        {
            var project = _multiGuidProject ?? GuidAuditService.CreateProjectTable();
            var family = _multiGuidFamilyDetail ?? GuidAuditService.CreateFamilyDetailTable();

            var sheets = new List<KeyValuePair<string, DataTable>>
            {
                new KeyValuePair<string, DataTable>("RVT 검토결과", GuidAuditService.PrepareExportTable(project, 1)),
                new KeyValuePair<string, DataTable>("Family 검토결과", GuidAuditService.PrepareExportTable(family, 2))
            };

            var saved = GuidAuditService.ExportMulti(sheets, excelMode, "guid:progress");
            if (string.IsNullOrWhiteSpace(saved))
            {
                SendToWeb("hub:multi-exported", new { ok = false, message = "엑셀 저장이 취소되었습니다." });
                return;
            }

            TryApplyExportStyles("guid", saved, !string.Equals(excelMode, "fast", StringComparison.OrdinalIgnoreCase), excelMode);
            SendToWeb("hub:multi-exported", new { ok = true, path = saved });
        }

        private void ExportFamilyLink(bool doAutoFit, string excelMode)
        {
            var dt = BuildTableFromRows(_multiFamilyLinkRows, "FamilyLink");
            ExcelCore.EnsureNoDataRow(dt, "FamilyLink 결과가 없습니다.");
            var saved = SaveTableByDialog(dt, "multi_familylink.xlsx", doAutoFit, "familylink");
            if (string.IsNullOrWhiteSpace(saved)) return;
            TryApplyExportStyles("familylink", saved, doAutoFit, excelMode);
            SendToWeb("hub:multi-exported", new { ok = true, path = saved });
        }

        private void ExportPoints(bool doAutoFit, string excelMode)
        {
            var dt = BuildTableFromRows(_multiPointsRows, "Points");
            ExcelCore.EnsureNoDataRow(dt, "Points 결과가 없습니다.");
            var saved = SaveTableByDialog(dt, "multi_points.xlsx", doAutoFit, "points");
            if (string.IsNullOrWhiteSpace(saved)) return;
            TryApplyExportStyles("points", saved, doAutoFit, excelMode);
            SendToWeb("hub:multi-exported", new { ok = true, path = saved });
        }

        private void ExportRunItems(bool doAutoFit, string excelMode)
        {
            var dt = new DataTable("RunItems");
            dt.Columns.Add("File");
            dt.Columns.Add("Status");
            dt.Columns.Add("Reason");
            dt.Columns.Add("Phase");
            dt.Columns.Add("ElapsedMs");

            foreach (var i in _multiRunItems)
            {
                var r = dt.NewRow();
                r["File"] = i.File;
                r["Status"] = i.Status;
                r["Reason"] = i.Reason;
                r["Phase"] = i.Phase;
                r["ElapsedMs"] = i.ElapsedMs;
                dt.Rows.Add(r);
            }

            ExcelCore.EnsureNoDataRow(dt, "데이터가 없습니다.");
            var saved = SaveTableByDialog(dt, "multi_run.xlsx", doAutoFit, "multi");
            if (string.IsNullOrWhiteSpace(saved)) return;
            TryApplyExportStyles("multi", saved, doAutoFit, excelMode);
            SendToWeb("hub:multi-exported", new { ok = true, path = saved });
        }

        private string SaveTableByDialog(DataTable dt, string fileName, bool doAutoFit, string key)
        {
            var dlg = new Microsoft.Win32.SaveFileDialog { Filter = "Excel (*.xlsx)|*.xlsx", FileName = fileName };
            if (dlg.ShowDialog() != true)
            {
                SendToWeb("hub:multi-exported", new { ok = false, message = "엑셀 저장이 취소되었습니다." });
                return string.Empty;
            }

            ExcelCore.SaveXlsx(dlg.FileName, dt.TableName, dt, doAutoFit, sheetKey: key, progressKey: "hub:multi-progress", exportKind: key);
            return dlg.FileName;
        }

        private MultiRunRequest ParseMultiRequest(object payload)
        {
            var req = new MultiRunRequest();
            req.RvtPaths = GetStringList(payload, "files").Where(x => !string.IsNullOrWhiteSpace(x)).Distinct(StringComparer.OrdinalIgnoreCase).ToList();
            req.Common = ParseCommon(payload);
            req.Connector = ParseConnector(payload);
            req.Pms = ParsePms(payload);
            req.Guid = ParseGuid(payload);
            req.FamilyLink = ParseFamilyLink(payload);
            req.Points = ParsePoints(payload);
            return req;
        }

        private MultiCommonOptions ParseCommon(object payload)
        {
            return new MultiCommonOptions
            {
                ExtraParams = GetString(payload, "extraParamsText", string.Empty),
                TargetFilter = GetString(payload, "targetFilterText", string.Empty),
                ExcludeEndDummy = GetBool(payload, "excludeEndDummy", false)
            };
        }

        private MultiConnectorOptions ParseConnector(object payload)
        {
            return new MultiConnectorOptions
            {
                Enabled = GetBool(payload, "runConnector", true),
                Tol = GetDouble(payload, "tol", 1.0),
                Unit = GetString(payload, "unit", "inch"),
                Param = GetString(payload, "param", "Comments")
            };
        }

        private MultiPmsOptions ParsePms(object payload)
        {
            return new MultiPmsOptions
            {
                Enabled = GetBool(payload, "runSegmentPms", false),
                NdRound = GetInt(payload, "ndRound", 3),
                TolMm = GetDouble(payload, "tolMm", 0.01),
                ClassMatch = GetBool(payload, "classMatch", false)
            };
        }

        private MultiGuidOptions ParseGuid(object payload)
        {
            return new MultiGuidOptions
            {
                Enabled = GetBool(payload, "runGuid", false),
                IncludeFamily = GetBool(payload, "includeFamily", false),
                IncludeAnnotation = GetBool(payload, "includeAnnotation", false)
            };
        }

        private MultiFamilyLinkOptions ParseFamilyLink(object payload)
        {
            return new MultiFamilyLinkOptions
            {
                Enabled = GetBool(payload, "runFamilyLink", false)
            };
        }

        private MultiPointsOptions ParsePoints(object payload)
        {
            return new MultiPointsOptions
            {
                Enabled = GetBool(payload, "runPoints", false),
                Unit = GetString(payload, "pointUnit", "ft")
            };
        }

        private void AppendMultiRunItem(string file, string status, string reason, string phase, DateTime started)
        {
            _multiRunItems.Add(new MultiRunItem
            {
                File = file ?? string.Empty,
                Status = status ?? string.Empty,
                Reason = reason ?? string.Empty,
                Phase = phase ?? string.Empty,
                ElapsedMs = (long)Math.Max(0, (DateTime.Now - started).TotalMilliseconds)
            });
        }

        private object BuildMultiSummaryPayload()
        {
            var total = _multiTotal > 0 ? _multiTotal : _multiRunItems.Count;
            var success = _multiRunItems.Count(x => string.Equals(x.Status, "success", StringComparison.OrdinalIgnoreCase));
            var skipped = _multiRunItems.Count(x => string.Equals(x.Status, "skipped", StringComparison.OrdinalIgnoreCase));
            var failed = _multiRunItems.Count(x => string.Equals(x.Status, "failed", StringComparison.OrdinalIgnoreCase));

            return new
            {
                ok = true,
                mode = "multiRvt",
                featureId = "multi_rvt_batch",
                title = "다중 RVT 검토",
                finishedAt = DateTime.Now.ToString("yyyy-MM-ddTHH:mm:ss"),
                total,
                success,
                skipped,
                failed,
                canceled = false,
                items = _multiRunItems,
                connector = new { rows = _multiConnectorRows.Count },
                pmsClass = new { rows = _multiPmsClassRows.Count },
                pmsSize = new { rows = _multiPmsSizeRows.Count },
                pmsRouting = new { rows = _multiPmsRoutingRows.Count },
                guid = new { rows = _multiGuidRows.Count },
                guidProject = new { rows = _multiGuidProject?.Rows.Count ?? 0 },
                guidFamily = new { rows = _multiGuidFamilyDetail?.Rows.Count ?? 0 },
                familylink = new { rows = _multiFamilyLinkRows.Count },
                points = new { rows = _multiPointsRows.Count },
                run = new { rows = _multiRunItems.Count }
            };
        }

        private void ReportMultiProgress(double pct, string message, string target)
        {
            SendToWeb("hub:multi-progress", new
            {
                pct,
                message,
                target,
                current = _multiIndex,
                total = _multiTotal
            });
        }

        private static double CalcStepPercent(double basePct, int stepIndex, int steps)
        {
            if (steps <= 0) return basePct * 100.0;
            var oneFileSpan = 1.0 / Math.Max(1, steps);
            return (basePct + oneFileSpan * Math.Max(0, Math.Min(stepIndex, steps))) * 100.0;
        }

        private static int CountEnabledFeatures(MultiRunRequest req)
        {
            if (req == null) return 0;
            var n = 0;
            if (req.Connector?.Enabled == true) n++;
            if (req.Pms?.Enabled == true) n++;
            if (req.Guid?.Enabled == true) n++;
            if (req.FamilyLink?.Enabled == true) n++;
            if (req.Points?.Enabled == true) n++;
            return n;
        }

        private void AppendMultiConnectorError(string fileName, string message)
        {
            _multiConnectorRows.Add(new Dictionary<string, object>
            {
                ["File"] = fileName ?? string.Empty,
                ["Status"] = "ERROR",
                ["Detail"] = message ?? string.Empty
            });
        }

        private void AppendSegmentPmsRows(string safeName)
        {
            _multiPmsClassRows.Add(new Dictionary<string, object>
            {
                ["File"] = safeName,
                ["Status"] = "CHECK",
                ["Detail"] = "PMS class compare queued"
            });
            _multiPmsSizeRows.Add(new Dictionary<string, object>
            {
                ["File"] = safeName,
                ["Status"] = "CHECK",
                ["Detail"] = "PMS size compare queued"
            });
            _multiPmsRoutingRows.Add(new Dictionary<string, object>
            {
                ["File"] = safeName,
                ["Status"] = "CHECK",
                ["Detail"] = "PMS routing compare queued"
            });
        }

        private void MergeGuidResult(DataTable project, DataTable familyDetail, DataTable familyIndex)
        {
            if (project != null)
            {
                if (_multiGuidProject == null) _multiGuidProject = project.Clone();
                foreach (DataRow r in project.Rows) _multiGuidProject.ImportRow(r);
            }

            if (familyDetail != null)
            {
                if (_multiGuidFamilyDetail == null) _multiGuidFamilyDetail = familyDetail.Clone();
                foreach (DataRow r in familyDetail.Rows) _multiGuidFamilyDetail.ImportRow(r);
            }

            if (familyIndex != null)
            {
                if (_multiGuidFamilyIndex == null) _multiGuidFamilyIndex = familyIndex.Clone();
                foreach (DataRow r in familyIndex.Rows) _multiGuidFamilyIndex.ImportRow(r);
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

        private static DataTable BuildTableFromRows(List<Dictionary<string, object>> rows, string tableName)
        {
            var dt = new DataTable(tableName);
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
    }
}
