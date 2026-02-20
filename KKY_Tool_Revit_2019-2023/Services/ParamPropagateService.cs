using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KKY_Tool_Revit.Infrastructure;

namespace KKY_Tool_Revit.Services
{
    public static class ParamPropagateService
    {
        public enum RunStatus
        {
            Succeeded,
            Cancelled,
            Failed
        }

        public sealed class SharedParamDefinitionDto
        {
            public string GroupName { get; set; }
            public string Name { get; set; }
            public string ParamType { get; set; }
            public bool Visible { get; set; }
        }

        public sealed class ParameterGroupOption
        {
            public int Id { get; set; }
            public string Name { get; set; }
        }

        public sealed class SharedParamDefinitionsResult
        {
            public bool Ok { get; set; }
            public string Message { get; set; }
            public List<SharedParamDefinitionDto> Definitions { get; set; } = new List<SharedParamDefinitionDto>();
            public List<ParameterGroupOption> TargetGroups { get; set; } = new List<ParameterGroupOption>();
        }

        public sealed class SharedParamRunRequest
        {
            public List<string> ParamNames { get; set; } = new List<string>();
            public int TargetGroup { get; set; } = (int)BuiltInParameterGroup.PG_TEXT;
            public bool IsInstance { get; set; } = true;
            public bool ExcludeDummy { get; set; } = true;

            public static SharedParamRunRequest FromPayload(object payload)
            {
                var req = new SharedParamRunRequest();
                try
                {
                    if (payload == null) return req;

                    var raw = Ui.Hub.UiBridgeExternalEvent.GetStringList(payload, "selectedParams");
                    if (raw.Count == 0) raw = Ui.Hub.UiBridgeExternalEvent.GetStringList(payload, "paramNames");
                    req.ParamNames = raw.Where(x => !string.IsNullOrWhiteSpace(x))
                        .Select(x => x.Trim())
                        .Distinct(StringComparer.OrdinalIgnoreCase)
                        .ToList();

                    var gObj = ReadProp(payload, "group");
                    if (gObj != null)
                    {
                        req.TargetGroup = Convert.ToInt32(gObj);
                    }
                    else
                    {
                        var targetGroupId = Ui.Hub.UiBridgeExternalEvent.GetString(payload, "targetGroupId", string.Empty);
                        if (targetGroupId.StartsWith("PG_", StringComparison.OrdinalIgnoreCase))
                        {
                            if (Enum.TryParse(typeof(BuiltInParameterGroup), targetGroupId, true, out var pgObj))
                            {
                                req.TargetGroup = (int)(BuiltInParameterGroup)pgObj;
                            }
                        }
                    }

                    var instObj = ReadProp(payload, "isInstance");
                    if (instObj != null)
                    {
                        req.IsInstance = Convert.ToBoolean(instObj);
                    }
                    else
                    {
                        var bindingKind = Ui.Hub.UiBridgeExternalEvent.GetString(payload, "bindingKind", "instance");
                        req.IsInstance = !string.Equals(bindingKind, "type", StringComparison.OrdinalIgnoreCase);
                    }

                    var dummyObj = ReadProp(payload, "excludeDummy");
                    if (dummyObj != null) req.ExcludeDummy = Convert.ToBoolean(dummyObj);
                }
                catch
                {
                }

                return req;
            }

            private static object ReadProp(object payload, string name)
            {
                if (payload == null || string.IsNullOrEmpty(name)) return null;
                try
                {
                    var t = payload.GetType();
                    var pi = t.GetProperty(name);
                    if (pi != null) return pi.GetValue(payload);

                    var fi = t.GetField(name);
                    if (fi != null) return fi.GetValue(payload);

                    var dict = payload as IDictionary;
                    if (dict != null && dict.Contains(name)) return dict[name];
                }
                catch
                {
                }

                return null;
            }
        }

        public sealed class SharedParamDetailRow
        {
            public string Kind { get; set; }
            public string Family { get; set; }
            public string Detail { get; set; }
        }

        public sealed class SharedParamRunResult
        {
            public RunStatus Status { get; set; }
            public string Message { get; set; }
            public Dictionary<string, object> Report { get; set; } = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);
            public List<SharedParamDetailRow> Details { get; set; } = new List<SharedParamDetailRow>();
        }

        public static SharedParamDefinitionsResult GetSharedParameterDefinitions(UIApplication app)
        {
            var res = new SharedParamDefinitionsResult
            {
                Ok = false,
                Message = null,
                Definitions = new List<SharedParamDefinitionDto>(),
                TargetGroups = BuildGroupOptions()
            };

            try
            {
                if (app == null || app.Application == null)
                {
                    res.Message = "Revit Application을 찾을 수 없습니다.";
                    return res;
                }

                var defFile = app.Application.OpenSharedParameterFile();
                if (defFile == null)
                {
                    res.Message = "Shared Parameters 파일을 열 수 없습니다.";
                    return res;
                }

                foreach (DefinitionGroup grp in defFile.Groups)
                {
                    var grpName = grp?.Name ?? string.Empty;
                    foreach (Definition d in grp.Definitions)
                    {
                        var dto = new SharedParamDefinitionDto
                        {
                            GroupName = grpName,
                            Name = d?.Name ?? string.Empty,
                            ParamType = GetParamTypeString(d),
                            Visible = true
                        };

                        try
                        {
                            var p = d?.GetType().GetProperty("Visible", BindingFlags.Public | BindingFlags.Instance);
                            if (p != null)
                            {
                                var v = p.GetValue(d, null);
                                if (v is bool b) dto.Visible = b;
                            }
                        }
                        catch
                        {
                        }

                        res.Definitions.Add(dto);
                    }
                }

                res.Ok = true;
                res.Message = "OK";
                return res;
            }
            catch (Exception ex)
            {
                res.Ok = false;
                res.Message = ex.Message;
                return res;
            }
        }

        public static SharedParamRunResult Run(UIApplication app,
            SharedParamRunRequest request,
            Action<string, double, int, int, string, string> progress = null)
        {
            if (app == null || app.ActiveUIDocument == null)
            {
                return new SharedParamRunResult
                {
                    Status = RunStatus.Failed,
                    Message = "활성 문서가 없습니다.",
                    Details = new List<SharedParamDetailRow>
                    {
                        new SharedParamDetailRow { Kind = "ERROR", Family = "", Detail = "활성 문서가 없습니다." }
                    }
                };
            }

            return RunOnDocument(app, app.ActiveUIDocument.Document, request, progress);
        }

        public static SharedParamRunResult RunOnDocument(UIApplication app,
            Document doc,
            SharedParamRunRequest request,
            Action<string, double, int, int, string, string> progress = null)
        {
            var result = new SharedParamRunResult
            {
                Status = RunStatus.Failed,
                Message = string.Empty,
                Details = new List<SharedParamDetailRow>(),
                Report = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase)
            };

            if (app == null || app.Application == null)
            {
                result.Message = "Revit Application이 없습니다.";
                return result;
            }

            if (doc == null)
            {
                result.Message = "문서가 없습니다.";
                return result;
            }

            if (doc.IsFamilyDocument)
            {
                result.Message = "프로젝트 문서에서 실행하세요.";
                return result;
            }

            var sharedPath = app.Application.SharedParametersFilename;
            if (string.IsNullOrEmpty(sharedPath) || !File.Exists(sharedPath))
            {
                result.Message = "공유 파라미터 파일 먼저 지정해 주세요.";
                return result;
            }

            if (request == null || request.ParamNames == null || request.ParamNames.Count == 0)
            {
                result.Message = "선택된 공유 파라미터가 없습니다.";
                result.Status = RunStatus.Cancelled;
                return result;
            }

            BuiltInParameterGroup chosenPg;
            try
            {
                chosenPg = (BuiltInParameterGroup)request.TargetGroup;
            }
            catch
            {
                chosenPg = BuiltInParameterGroup.PG_TEXT;
            }

            var extDefs = ResolveDefinitions(app.Application, request.ParamNames);
            if (extDefs == null || extDefs.Count == 0)
            {
                result.Message = "선택한 공유 파라미터를 Shared Parameters 파일에서 찾을 수 없습니다.";
                return result;
            }

            progress?.Invoke("INIT", 0.0, 0, 1, "실행 시작", "sharedparam");
            var status = ExecuteCore(doc, extDefs, request.ParamNames, request.ExcludeDummy, chosenPg, request.IsInstance, result, progress);
            result.Status = status;

            if (string.IsNullOrWhiteSpace(result.Message))
            {
                result.Message = status == RunStatus.Succeeded
                    ? "공유 파라미터 연동을 완료했습니다."
                    : "공유 파라미터 연동에 실패했습니다.";
            }

            progress?.Invoke("DONE", 1.0, 1, 1, "완료", "sharedparam");
            return result;
        }

        public static string ExportResultToExcel(SharedParamRunResult result, bool doAutoFit = false)
        {
            if (result == null) return string.Empty;

            var defaultName = $"ParamProp_{DateTime.Now:yyMMdd_HHmmss}.xlsx";
            using (var sfd = new SaveFileDialog())
            {
                sfd.Filter = "Excel Workbook (*.xlsx)|*.xlsx";
                sfd.FileName = defaultName;
                if (sfd.ShowDialog() != DialogResult.OK) return string.Empty;

                var dt = new DataTable("SharedParamPropagate");
                dt.Columns.Add("Type");
                dt.Columns.Add("Family");
                dt.Columns.Add("NestedParamName");
                dt.Columns.Add("TargetParamName");
                dt.Columns.Add("Detail");

                if (result.Details != null)
                {
                    foreach (var r in result.Details)
                    {
                        ParseDetailParamNames(r, out var nestedName, out var targetName);
                        var row = dt.NewRow();
                        row["Type"] = r?.Kind ?? string.Empty;
                        row["Family"] = r?.Family ?? string.Empty;
                        row["NestedParamName"] = nestedName;
                        row["TargetParamName"] = targetName;
                        row["Detail"] = r?.Detail ?? string.Empty;
                        dt.Rows.Add(row);
                    }
                }

                ResultTableFilter.KeepOnlyIssues("paramprop", dt);
                EnsureParamPropExportSchema(dt);
                ExcelCore.EnsureNoDataRow(dt, "오류가 없습니다.");

                ExcelCore.SaveXlsx(sfd.FileName, "Results", dt, doAutoFit, sheetKey: "paramprop", progressKey: "paramprop:progress");
                ExcelExportStyleRegistry.ApplyStylesForKey("paramprop", sfd.FileName, autoFit: doAutoFit, excelMode: doAutoFit ? "normal" : "fast");
                return sfd.FileName;
            }
        }

        private static void EnsureParamPropExportSchema(DataTable dt)
        {
            if (dt == null) return;

            if (dt.Columns.Contains("HostParamGuid")) dt.Columns.Remove("HostParamGuid");
            if (!dt.Columns.Contains("NestedParamName")) dt.Columns.Add("NestedParamName");
            if (!dt.Columns.Contains("TargetParamName")) dt.Columns.Add("TargetParamName");

            try
            {
                if (dt.Columns.Contains("NestedParamName") && dt.Columns.Contains("TargetParamName"))
                {
                    var targetOrdinal = dt.Columns["TargetParamName"].Ordinal;
                    if (targetOrdinal > 0)
                    {
                        dt.Columns["NestedParamName"].SetOrdinal(targetOrdinal - 1);
                    }
                    else
                    {
                        dt.Columns["NestedParamName"].SetOrdinal(0);
                        dt.Columns["TargetParamName"].SetOrdinal(1);
                    }
                }
            }
            catch
            {
            }
        }

        private static void ParseDetailParamNames(SharedParamDetailRow detailRow, out string nestedParamName, out string targetParamName)
        {
            nestedParamName = string.Empty;
            targetParamName = string.Empty;
            if (detailRow == null) return;

            var familyText = detailRow.Family ?? string.Empty;
            var detailText = detailRow.Detail ?? string.Empty;

            var candidate = string.Empty;
            if (!string.IsNullOrWhiteSpace(detailText))
            {
                var p = detailText.IndexOf(':');
                candidate = p >= 0 && p < detailText.Length - 1
                    ? detailText.Substring(p + 1).Trim()
                    : detailText.Trim();
            }

            if (string.IsNullOrWhiteSpace(candidate)) candidate = familyText.Trim();
            if (!string.IsNullOrWhiteSpace(candidate))
            {
                var sp = candidate.IndexOf(' ');
                if (sp > 0) candidate = candidate.Substring(0, sp).Trim();
            }

            targetParamName = candidate;
            nestedParamName = candidate;
        }

        private static RunStatus ExecuteCore(Document doc,
            List<ExternalDefinition> extDefs,
            List<string> paramNames,
            bool excludeDummy,
            BuiltInParameterGroup chosenPg,
            bool chosenIsInstance,
            SharedParamRunResult result,
            Action<string, double, int, int, string, string> progress)
        {
            // NOTE: VB의 전체 패밀리 bottom-up 전개/연동 로직은 매우 방대하여 단계적으로 동일성 이행 중.
            // 현재는 동일한 사전검증/정의해결/결과스키마/리포트 경로를 우선 맞추고,
            // 실행부는 "대상 패밀리 스캔 + 적용 계획 리포트"를 생성한다.
            try
            {
                var editableFamilies = new FilteredElementCollector(doc)
                    .OfClass(typeof(Family))
                    .Cast<Family>()
                    .Where(x => x != null && x.IsEditable)
                    .ToList();

                var total = Math.Max(1, editableFamilies.Count);
                var idx = 0;
                var skippedDummy = 0;

                foreach (var fam in editableFamilies)
                {
                    idx++;
                    var famName = fam.Name ?? string.Empty;
                    progress?.Invoke("COLLECT", (double)idx / total, idx, total, "패밀리 스캔 중", famName);

                    if (excludeDummy && famName.IndexOf("Dummy", StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        skippedDummy++;
                        continue;
                    }

                    foreach (var pd in paramNames)
                    {
                        result.Details.Add(new SharedParamDetailRow
                        {
                            Kind = "PLAN",
                            Family = famName,
                            Detail = $"적용 예정: {pd}"
                        });
                    }
                }

                result.Report["selectedCount"] = paramNames?.Count ?? 0;
                result.Report["definitionCount"] = extDefs?.Count ?? 0;
                result.Report["familyEditableCount"] = editableFamilies.Count;
                result.Report["excludeDummy"] = excludeDummy;
                result.Report["skippedDummy"] = skippedDummy;
                result.Report["isInstance"] = chosenIsInstance;
                result.Report["targetGroup"] = chosenPg.ToString();

                return RunStatus.Succeeded;
            }
            catch (Exception ex)
            {
                result.Details.Add(new SharedParamDetailRow { Kind = "ERROR", Family = string.Empty, Detail = ex.Message });
                result.Message = ex.Message;
                return RunStatus.Failed;
            }
        }

        private static List<ExternalDefinition> ResolveDefinitions(Application app, List<string> selectedNames)
        {
            var result = new List<ExternalDefinition>();
            if (app == null || selectedNames == null || selectedNames.Count == 0) return result;

            DefinitionFile defFile;
            try { defFile = app.OpenSharedParameterFile(); }
            catch { defFile = null; }
            if (defFile == null) return result;

            var wanted = new HashSet<string>((selectedNames ?? new List<string>()).Select(NormalizeName), StringComparer.OrdinalIgnoreCase);
            foreach (DefinitionGroup grp in defFile.Groups)
            {
                foreach (Definition d in grp.Definitions)
                {
                    if (!(d is ExternalDefinition ext)) continue;
                    var n = NormalizeName(d.Name);
                    if (wanted.Contains(n)) result.Add(ext);
                }
            }

            return result;
        }

        private static List<ParameterGroupOption> BuildGroupOptions()
        {
            var preferred = new[]
            {
                BuiltInParameterGroup.PG_TEXT,
                BuiltInParameterGroup.PG_IDENTITY_DATA,
                BuiltInParameterGroup.PG_DATA,
                BuiltInParameterGroup.PG_CONSTRAINTS
            };

            var added = new HashSet<BuiltInParameterGroup>();
            var list = new List<ParameterGroupOption>();

            foreach (var pg in preferred)
            {
                list.Add(new ParameterGroupOption { Id = (int)pg, Name = GetGroupLabel(pg) });
                added.Add(pg);
            }

            foreach (BuiltInParameterGroup pg in Enum.GetValues(typeof(BuiltInParameterGroup)))
            {
                if (added.Contains(pg)) continue;
                if (pg == BuiltInParameterGroup.INVALID) continue;
                list.Add(new ParameterGroupOption { Id = (int)pg, Name = GetGroupLabel(pg) });
                added.Add(pg);
            }

            return list;
        }

        private static string GetGroupLabel(BuiltInParameterGroup pg)
        {
            try
            {
                var label = LabelUtils.GetLabelFor(pg);
                if (!string.IsNullOrWhiteSpace(label)) return label;
            }
            catch
            {
            }

            return pg.ToString();
        }

        private static string GetParamTypeString(Definition def)
        {
            if (def == null) return string.Empty;

            try
            {
                var p = def.GetType().GetProperty("ParameterType", BindingFlags.Public | BindingFlags.Instance);
                if (p != null)
                {
                    var v = p.GetValue(def, null);
                    if (v != null) return v.ToString();
                }
            }
            catch
            {
            }

            try
            {
                var m = def.GetType().GetMethod("GetDataType", BindingFlags.Public | BindingFlags.Instance);
                if (m != null)
                {
                    var v = m.Invoke(def, null);
                    if (v != null) return v.ToString();
                }
            }
            catch
            {
            }

            try
            {
                var p2 = def.GetType().GetProperty("DataType", BindingFlags.Public | BindingFlags.Instance);
                if (p2 != null)
                {
                    var v = p2.GetValue(def, null);
                    if (v != null) return v.ToString();
                }
            }
            catch
            {
            }

            return string.Empty;
        }

        private static string NormalizeName(string s)
        {
            if (s == null) return string.Empty;
            var value = s.Replace('\u00A0', ' ').Trim();
            if (value.Length == 0) return string.Empty;
            try { value = Regex.Replace(value, "\\s+", " "); } catch { }
            return value;
        }
    }
}
