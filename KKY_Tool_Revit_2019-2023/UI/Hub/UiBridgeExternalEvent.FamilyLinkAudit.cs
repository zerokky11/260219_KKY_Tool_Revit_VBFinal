using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Windows.Forms;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KKY_Tool_Revit.Exports;
using KKY_Tool_Revit.Services;

namespace KKY_Tool_Revit.UI.Hub
{
    public sealed partial class UiBridgeExternalEvent
    {
        private List<FamilyLinkAuditRow> _familyLinkLastRows;

        private void HandleFamilyLinkInit(UIApplication app)
        {
            try
            {
                var sourcePath = string.Empty;
                try
                {
                    sourcePath = app.Application.SharedParametersFilename;
                }
                catch
                {
                    // ignore
                }

                if (string.IsNullOrWhiteSpace(sourcePath))
                {
                    SendToWeb("familylink:error", new
                    {
                        message = "Shared Parameters 파일이 설정되어 있지 않습니다.",
                        detail = "Revit 옵션에서 Shared Parameters 파일 경로를 설정하세요."
                    });
                    return;
                }

                DefinitionFile defFile = null;
                try
                {
                    defFile = app.Application.OpenSharedParameterFile();
                }
                catch (Exception ex)
                {
                    SendToWeb("familylink:error", new
                    {
                        message = "Shared Parameters 파일을 열 수 없습니다.",
                        detail = ex.Message
                    });
                    return;
                }

                if (defFile == null)
                {
                    SendToWeb("familylink:error", new
                    {
                        message = "Shared Parameters 파일을 읽지 못했습니다.",
                        detail = sourcePath
                    });
                    return;
                }

                var items = new List<object>();
                foreach (DefinitionGroup g in defFile.Groups)
                {
                    if (g == null) continue;
                    foreach (Definition def in g.Definitions)
                    {
                        var ext = def as ExternalDefinition;
                        if (ext == null) continue;
                        items.Add(new
                        {
                            name = ext.Name,
                            guid = ext.GUID.ToString("D"),
                            groupName = g.Name,
                            dataTypeToken = SafeDefTypeToken(ext)
                        });
                    }
                }

                SendToWeb("familylink:sharedparams", new
                {
                    sourcePath,
                    items
                });
            }
            catch (Exception ex)
            {
                SendToWeb("familylink:error", new
                {
                    message = "Shared Parameters 목록 로드 실패",
                    detail = ex.Message
                });
            }
        }

        private void HandleFamilyLinkPickRvts()
        {
            using (var dlg = new OpenFileDialog())
            {
                dlg.Filter = "Revit Project (*.rvt)|*.rvt";
                dlg.Multiselect = true;
                dlg.Title = "패밀리 연동 검토 대상 RVT 선택";
                dlg.RestoreDirectory = true;

                if (dlg.ShowDialog() != DialogResult.OK) return;

                var files = new List<string>();
                var dedup = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                foreach (var p in dlg.FileNames)
                {
                    if (string.IsNullOrWhiteSpace(p)) continue;
                    if (dedup.Add(p)) files.Add(p);
                }

                SendToWeb("familylink:rvts-picked", new { paths = files });
            }
        }

        private void HandleFamilyLinkRun(UIApplication app, object payload)
        {
            _familyLinkLastRows = null;

            var rvtPaths = ExtractStringList(payload, "rvtPaths");
            var targets = ExtractTargets(payload);

            if (rvtPaths.Count == 0)
            {
                SendToWeb("familylink:error", new
                {
                    message = "검토할 RVT 파일이 없습니다.",
                    detail = "RVT 목록을 추가하세요."
                });
                return;
            }

            if (targets.Count == 0)
            {
                SendToWeb("familylink:error", new
                {
                    message = "검토할 파라미터가 없습니다.",
                    detail = "Shared Parameters 목록에서 대상 파라미터를 선택하세요."
                });
                return;
            }

            try
            {
                var rows = FamilyLinkAuditService.Run(app, rvtPaths, targets, ReportFamilyLinkProgress);
                var filteredRows = FilterFamilyLinkIssueRows(rows);
                _familyLinkLastRows = filteredRows;

                var schema = FamilyLinkAuditExport.Schema;
                var payloadRows = filteredRows.Select(ToRowDict).ToList();

                SendToWeb("familylink:result", new
                {
                    schema,
                    rows = payloadRows
                });
            }
            catch (Exception ex)
            {
                SendToWeb("familylink:error", new
                {
                    message = "패밀리 연동 검토 실행 실패",
                    detail = ex.Message
                });
            }
        }

        private List<FamilyLinkAuditRow> FilterFamilyLinkIssueRows(List<FamilyLinkAuditRow> rows)
        {
            if (rows == null) return new List<FamilyLinkAuditRow>();

            var table = FamilyLinkAuditExport.ToDataTable(rows);
            var filteredTable = FilterIssueRowsCopy("familylink", table);

            var result = new List<FamilyLinkAuditRow>();
            if (filteredTable == null) return result;

            foreach (DataRow dr in filteredTable.Rows)
            {
                result.Add(new FamilyLinkAuditRow
                {
                    FileName = SafeStr(Convert.ToString(dr["FileName"])),
                    HostFamilyName = SafeStr(Convert.ToString(dr["HostFamilyName"])),
                    HostFamilyCategory = SafeStr(Convert.ToString(dr["HostFamilyCategory"])),
                    NestedFamilyName = SafeStr(Convert.ToString(dr["NestedFamilyName"])),
                    NestedTypeName = SafeStr(Convert.ToString(dr["NestedTypeName"])),
                    NestedCategory = SafeStr(Convert.ToString(dr["NestedCategory"])),
                    NestedParamName = SafeStr(Convert.ToString(dr["NestedParamName"])),
                    TargetParamName = SafeStr(Convert.ToString(dr["TargetParamName"])),
                    ExpectedGuid = SafeStr(Convert.ToString(dr["ExpectedGuid"])),
                    FoundScope = SafeStr(Convert.ToString(dr["FoundScope"])),
                    NestedParamGuid = SafeStr(Convert.ToString(dr["NestedParamGuid"])),
                    NestedParamDataType = SafeStr(Convert.ToString(dr["NestedParamDataType"])),
                    AssocHostParamName = SafeStr(Convert.ToString(dr["AssocHostParamName"])),
                    HostParamGuid = SafeStr(Convert.ToString(dr["HostParamGuid"])),
                    HostParamIsShared = SafeStr(Convert.ToString(dr["HostParamIsShared"])),
                    Issue = SafeStr(Convert.ToString(dr["Issue"])),
                    Notes = SafeStr(Convert.ToString(dr["Notes"]))
                });
            }

            return result;
        }

        private void HandleFamilyLinkExport(object payload)
        {
            if (_familyLinkLastRows == null)
            {
                SendToWeb("familylink:error", new
                {
                    message = "내보낼 결과가 없습니다.",
                    detail = "먼저 스캔을 실행하세요."
                });
                return;
            }

            var fastExport = ExtractBool(payload, "fastExport", true);
            var autoFit = ParseExcelMode(payload);
            if (ExtractBool(payload, "fastExport", false)) autoFit = false;
            fastExport = !autoFit;

            try
            {
                var savedPath = FamilyLinkAuditExport.Export(_familyLinkLastRows, fastExport, autoFit);
                if (string.IsNullOrWhiteSpace(savedPath))
                {
                    SendToWeb("familylink:exported", new
                    {
                        ok = false,
                        message = "엑셀/CSV 내보내기가 취소되었습니다."
                    });
                    return;
                }

                SendToWeb("familylink:exported", new
                {
                    ok = true,
                    path = savedPath
                });
            }
            catch (Exception ex)
            {
                SendToWeb("familylink:exported", new
                {
                    ok = false,
                    message = ex.Message
                });
                SendToWeb("familylink:error", new
                {
                    message = "엑셀/CSV 내보내기 실패",
                    detail = ex.Message
                });
            }
        }

        private void ReportFamilyLinkProgress(int percent, string message)
        {
            try
            {
                SendToWeb("familylink:progress", new
                {
                    percent = Math.Max(0, Math.Min(100, percent)),
                    message = message ?? ""
                });
            }
            catch
            {
                // ignore
            }
        }

        private List<string> ExtractStringList(object payload, string key)
        {
            var res = new List<string>();
            var payloadValue = GetPropFamilyLink(payload, key);
            if (payloadValue == null) return res;

            var payloadItems = payloadValue as IEnumerable;
            if (payloadItems == null || payloadValue is string)
            {
                var singleValue = payloadValue as string;
                if (!string.IsNullOrWhiteSpace(singleValue))
                {
                    res.Add(singleValue);
                }
                return res;
            }

            foreach (var o in payloadItems)
            {
                if (o == null) continue;
                var s = o.ToString();
                if (!string.IsNullOrWhiteSpace(s)) res.Add(s);
            }

            return res;
        }

        private List<FamilyLinkTargetParam> ExtractTargets(object payload)
        {
            var list = new List<FamilyLinkTargetParam>();
            var payloadValue = GetPropFamilyLink(payload, "targets");
            var payloadItems = payloadValue as IEnumerable;
            if (payloadItems == null || payloadValue is string) return list;

            foreach (var o in payloadItems)
            {
                var name = GetPropFamilyLink(o, "name") as string;
                var guidStr = GetPropFamilyLink(o, "guid") as string;
                if (string.IsNullOrWhiteSpace(name) || string.IsNullOrWhiteSpace(guidStr)) continue;
                Guid g;
                if (!Guid.TryParse(guidStr, out g)) continue;

                var item = new FamilyLinkTargetParam
                {
                    Name = name.Trim(),
                    Guid = g
                };
                list.Add(item);
            }

            return list;
        }

        private Dictionary<string, object> ToRowDict(FamilyLinkAuditRow row)
        {
            var d = new Dictionary<string, object>(StringComparer.Ordinal)
            {
                ["FileName"] = row.FileName,
                ["HostFamilyName"] = row.HostFamilyName,
                ["HostFamilyCategory"] = row.HostFamilyCategory,
                ["NestedFamilyName"] = row.NestedFamilyName,
                ["NestedTypeName"] = row.NestedTypeName,
                ["NestedCategory"] = row.NestedCategory,
                ["TargetParamName"] = row.TargetParamName,
                ["ExpectedGuid"] = row.ExpectedGuid,
                ["FoundScope"] = row.FoundScope,
                ["NestedParamGuid"] = row.NestedParamGuid,
                ["NestedParamDataType"] = row.NestedParamDataType,
                ["AssocHostParamName"] = row.AssocHostParamName,
                ["HostParamGuid"] = row.HostParamGuid,
                ["HostParamIsShared"] = row.HostParamIsShared,
                ["Issue"] = row.Issue,
                ["Notes"] = row.Notes
            };
            return d;
        }

        private bool ExtractBool(object payload, string key, bool defaultValue)
        {
            try
            {
                if (payload == null) return defaultValue;
                var raw = GetPropFamilyLink(payload, key);
                if (raw == null) return defaultValue;
                return Convert.ToBoolean(raw);
            }
            catch
            {
                return defaultValue;
            }
        }

        private static string SafeDefTypeToken(Definition defn)
        {
            if (defn == null) return "";
            try
            {
                return SafeStr(defn.ParameterType.ToString());
            }
            catch
            {
                return "";
            }
        }

        private static string SafeStr(string s)
        {
            return s ?? "";
        }

        private static object GetPropFamilyLink(object obj, string prop)
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
            var p = t.GetProperty(prop, System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.IgnoreCase);
            if (p == null) return null;
            return p.GetValue(obj, null);
        }
    }
}
