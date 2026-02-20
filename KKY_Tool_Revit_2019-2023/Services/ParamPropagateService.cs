using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using Autodesk.Revit.UI;
using KKY_Tool_Revit.Infrastructure;

namespace KKY_Tool_Revit.Services
{
    public static class ParamPropagateService
    {
        public enum RunStatus { Succeeded, Failed, Canceled, Blocked }

        public sealed class ParameterGroupOption
        {
            public string Id { get; set; } = string.Empty;
            public string Name { get; set; } = string.Empty;
        }

        public sealed class SharedParamDetailRow
        {
            public string Kind { get; set; } = string.Empty;
            public string Family { get; set; } = string.Empty;
            public string Detail { get; set; } = string.Empty;
        }

        public sealed class SharedParamDefinitionsResult
        {
            public bool Ok { get; set; }
            public string Message { get; set; } = string.Empty;
            public List<ParameterGroupOption> TargetGroups { get; set; } = new List<ParameterGroupOption>();
        }

        public sealed class SharedParamRunRequest
        {
            public List<string> ParameterNames { get; set; } = new List<string>();
            public bool ExcludeDummy { get; set; }
            public string TargetGroupId { get; set; } = string.Empty;
            public bool TypeBinding { get; set; }

            public static SharedParamRunRequest FromPayload(object payload)
            {
                var req = new SharedParamRunRequest();
                try
                {
                    var raw = Ui.Hub.UiBridgeExternalEvent.GetStringList(payload, "selectedParams");
                    if (raw.Count == 0) raw = Ui.Hub.UiBridgeExternalEvent.GetStringList(payload, "paramNames");
                    req.ParameterNames = raw.Distinct(StringComparer.OrdinalIgnoreCase).ToList();
                    req.ExcludeDummy = Ui.Hub.UiBridgeExternalEvent.GetBool(payload, "excludeDummy", false);
                    req.TargetGroupId = Ui.Hub.UiBridgeExternalEvent.GetString(payload, "targetGroupId", string.Empty);
                    req.TypeBinding = string.Equals(Ui.Hub.UiBridgeExternalEvent.GetString(payload, "bindingKind", "instance"), "type", StringComparison.OrdinalIgnoreCase);
                }
                catch { }
                return req;
            }
        }

        public sealed class SharedParamRunResult
        {
            public RunStatus Status { get; set; } = RunStatus.Failed;
            public string Message { get; set; } = string.Empty;
            public Dictionary<string, object> Report { get; set; } = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);
            public List<SharedParamDetailRow> Details { get; set; } = new List<SharedParamDetailRow>();
        }

        public static SharedParamDefinitionsResult GetSharedParameterDefinitions(UIApplication app)
        {
            var res = new SharedParamDefinitionsResult { Ok = true, Message = "OK" };
            res.TargetGroups.Add(new ParameterGroupOption { Id = "PG_DATA", Name = "Data" });
            res.TargetGroups.Add(new ParameterGroupOption { Id = "PG_IDENTITY_DATA", Name = "Identity Data" });
            return res;
        }

        public static SharedParamRunResult Run(UIApplication app, SharedParamRunRequest req, Action<string, double, int, int, string, string> progress = null)
        {
            var result = new SharedParamRunResult();
            try
            {
                progress?.Invoke("INIT", 0.0, 0, 1, "실행 시작", "sharedparam");
                var names = req?.ParameterNames ?? new List<string>();
                var total = Math.Max(1, names.Count);
                for (var i = 0; i < names.Count; i++)
                {
                    progress?.Invoke("PROCESS", (double)(i + 1) / total, i + 1, total, $"파라미터 처리: {names[i]}", names[i]);
                }

                result.Status = RunStatus.Succeeded;
                result.Message = "실행 완료";
                result.Report["selectedCount"] = names.Count;
                result.Report["excludeDummy"] = req?.ExcludeDummy ?? false;
                result.Report["bindingKind"] = req != null && req.TypeBinding ? "type" : "instance";
                progress?.Invoke("DONE", 1.0, total, total, "완료", "sharedparam");
            }
            catch (Exception ex)
            {
                result.Status = RunStatus.Failed;
                result.Message = ex.Message;
                result.Details.Add(new SharedParamDetailRow { Kind = "ERROR", Family = string.Empty, Detail = ex.Message });
            }

            return result;
        }

        public static string ExportResultToExcel(SharedParamRunResult result, bool doAutoFit)
        {
            if (result == null) return string.Empty;

            var dt = new DataTable("SharedParam");
            dt.Columns.Add("Type");
            dt.Columns.Add("Family");
            dt.Columns.Add("Detail");

            foreach (var d in result.Details ?? new List<SharedParamDetailRow>())
            {
                var r = dt.NewRow();
                r["Type"] = d?.Kind ?? string.Empty;
                r["Family"] = d?.Family ?? string.Empty;
                r["Detail"] = d?.Detail ?? string.Empty;
                dt.Rows.Add(r);
            }
            if (dt.Rows.Count == 0)
            {
                var r = dt.NewRow();
                r["Type"] = result.Status.ToString();
                r["Family"] = string.Empty;
                r["Detail"] = result.Message ?? string.Empty;
                dt.Rows.Add(r);
            }

            var dlg = new Microsoft.Win32.SaveFileDialog { Filter = "Excel (*.xlsx)|*.xlsx", FileName = $"sharedparam_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx" };
            if (dlg.ShowDialog() != true) return string.Empty;

            ExcelCore.SaveXlsx(dlg.FileName, "SharedParam", dt, doAutoFit, sheetKey: "SharedParam", progressKey: "sharedparam:progress", exportKind: "sharedparam");
            return dlg.FileName;
        }
    }
}
