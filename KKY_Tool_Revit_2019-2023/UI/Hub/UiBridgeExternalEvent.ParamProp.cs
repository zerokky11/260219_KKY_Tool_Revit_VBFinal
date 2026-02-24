using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using Autodesk.Revit.UI;
using KKY_Tool_Revit.Services;

namespace KKY_Tool_Revit.UI.Hub
{
    public sealed partial class UiBridgeExternalEvent
    {
        private static ParamPropagateService.SharedParamRunResult _lastParamResult;

        private void HandleSharedParamList(UIApplication app, object payload)
        {
            try
            {
                var status = SharedParameterStatusService.GetStatus(app);
                if (status == null || !string.Equals(status.Status, "ok", StringComparison.OrdinalIgnoreCase))
                {
                    var msg = string.IsNullOrWhiteSpace(status?.WarningMessage)
                        ? "Shared Parameter 파일 상태가 올바르지 않습니다."
                        : status.WarningMessage;
                    SendToWeb("sharedparam:list", new { ok = false, message = msg, items = new List<object>() });
                    return;
                }

                var res = ParamPropagateService.GetSharedParameterDefinitions(app);
                var items = SharedParameterStatusService.ListDefinitions(app).Select(d => new
                {
                    name = d.Name,
                    guid = d.Guid,
                    groupName = d.GroupName,
                    dataTypeToken = d.DataTypeToken
                }).ToList();

                var shaped = new
                {
                    ok = res != null && res.Ok,
                    message = res == null ? null : res.Message,
                    items,
                    definitions = items.Select(d => new
                    {
                        groupName = d.groupName,
                        name = d.name,
                        paramType = d.dataTypeToken,
                        visible = true
                    }).ToList(),
                    targetGroups = (res?.TargetGroups ?? new List<ParamPropagateService.ParameterGroupOption>()).Select(g => new
                    {
                        id = g.Id,
                        name = g.Name
                    }).ToList()
                };
                SendToWeb("sharedparam:list", shaped);
            }
            catch (Exception ex)
            {
                SendToWeb("sharedparam:list", new { ok = false, message = ex.Message, items = new List<object>() });
                SendToWeb("revit:error", new { message = ex.Message });
            }
        }

        private void HandleSharedParamStatus(UIApplication app, object payload)
        {
            try
            {
                var status = SharedParameterStatusService.GetStatus(app);
                var shaped = new
                {
                    path = status.Path,
                    isSet = status.IsSet,
                    existsOnDisk = status.ExistsOnDisk,
                    canOpen = status.CanOpen,
                    status = status.Status,
                    statusLabel = status.StatusLabel,
                    warning = status.WarningMessage,
                    errorMessage = status.ErrorMessage
                };
                SendToWeb("sharedparam:status", shaped);
            }
            catch (Exception ex)
            {
                SendToWeb("sharedparam:status", new
                {
                    status = "error",
                    statusLabel = "조회 실패",
                    warning = "Shared Parameter 상태 조회에 실패했습니다.",
                    errorMessage = ex.Message
                });
            }
        }

        private void HandleSharedParamRun(UIApplication app, object payload)
        {
            try
            {
                var sharedStatus = SharedParameterStatusService.GetStatus(app);
                if (sharedStatus == null || !string.Equals(sharedStatus.Status, "ok", StringComparison.OrdinalIgnoreCase))
                {
                    var msg = string.IsNullOrWhiteSpace(sharedStatus?.WarningMessage)
                        ? "Shared Parameter 파일 상태가 올바르지 않습니다."
                        : sharedStatus.WarningMessage;
                    SendToWeb("sharedparam:done", new { ok = false, status = "blocked", message = msg });
                    SendToWeb("paramprop:done", new { ok = false, status = "blocked", message = msg });
                    SendToWeb("revit:error", new { message = "공유 파라미터 연동 실패: " + msg });
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
                    return;
                }

                var req = ParamPropagateService.SharedParamRunRequest.FromPayload(payload);
                var res = ParamPropagateService.Run(app, req, ReportParamPropProgress);
                _lastParamResult = res;

                var status = res == null ? ParamPropagateService.RunStatus.Failed : res.Status;
                var ok = status == ParamPropagateService.RunStatus.Succeeded;

                var responsePayload = new Dictionary<string, object>();
                responsePayload["ok"] = ok;
                responsePayload["status"] = status.ToString().ToLowerInvariant();
                responsePayload["message"] = res != null ? res.Message : null;

                if (res != null)
                {
                    var filteredDetails = FilterParamPropIssueDetails(res.Details);
                    responsePayload["report"] = res.Report;
                    responsePayload["details"] = filteredDetails.Select(d => new
                    {
                        kind = d.Kind,
                        family = d.Family,
                        detail = d.Detail
                    }).ToList();
                }

                SendToWeb("sharedparam:done", responsePayload);
                SendToWeb("paramprop:done", responsePayload);

                if (!ok)
                {
                    var msg = res == null ? "실패" : res.Message;
                    SendToWeb("revit:error", new { message = "공유 파라미터 연동 실패: " + msg });
                }
            }
            catch (Exception ex)
            {
                SendToWeb("sharedparam:done", new { ok = false, status = "failed", message = ex.Message });
                SendToWeb("paramprop:done", new { ok = false, status = "failed", message = ex.Message });
                SendToWeb("revit:error", new { message = "공유 파라미터 연동 실패: " + ex.Message });
            }
        }

        private List<ParamPropagateService.SharedParamDetailRow> FilterParamPropIssueDetails(List<ParamPropagateService.SharedParamDetailRow> details)
        {
            var src = details ?? new List<ParamPropagateService.SharedParamDetailRow>();

            var dt = new DataTable("ParamPropUi");
            dt.Columns.Add("Type");
            dt.Columns.Add("Family");
            dt.Columns.Add("Detail");

            foreach (var d in src)
            {
                var row = dt.NewRow();
                row["Type"] = d == null ? "" : (d.Kind ?? "");
                row["Family"] = d == null ? "" : (d.Family ?? "");
                row["Detail"] = d == null ? "" : (d.Detail ?? "");
                dt.Rows.Add(row);
            }

            var filtered = FilterIssueRowsCopy("paramprop", dt);
            var result = new List<ParamPropagateService.SharedParamDetailRow>();
            if (filtered == null) return result;

            foreach (DataRow r in filtered.Rows)
            {
                result.Add(new ParamPropagateService.SharedParamDetailRow
                {
                    Kind = Convert.ToString(r["Type"]),
                    Family = Convert.ToString(r["Family"]),
                    Detail = Convert.ToString(r["Detail"])
                });
            }

            return result;
        }

        private void ReportParamPropProgress(string phase, double phaseProgress, int current, int total, string message, string target)
        {
            SendToWeb("paramprop:progress", new
            {
                phase,
                phaseProgress,
                current,
                total,
                message,
                target
            });
        }

        private void HandleSharedParamExport(UIApplication app, object payload)
        {
            try
            {
                if (_lastParamResult == null)
                {
                    SendToWeb("sharedparam:exported", new { ok = false, message = "최근 실행 결과가 없습니다." });
                    return;
                }

                var doAutoFit = ParseExcelMode(payload);
                var saved = ParamPropagateService.ExportResultToExcel(_lastParamResult, doAutoFit);
                if (string.IsNullOrWhiteSpace(saved))
                {
                    SendToWeb("sharedparam:exported", new { ok = false, message = "엑셀 내보내기가 취소되었습니다." });
                    return;
                }

                SendToWeb("sharedparam:exported", new { ok = true, path = saved });
            }
            catch (Exception ex)
            {
                SendToWeb("sharedparam:exported", new { ok = false, message = ex.Message });
                SendToWeb("revit:error", new { message = "엑셀 내보내기 실패: " + ex.Message });
            }
        }
    }
}
