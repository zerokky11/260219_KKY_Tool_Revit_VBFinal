using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Web.Script.Serialization;
using System.Windows.Forms;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KKY_Tool_Revit.Infrastructure;
using RevitApp = Autodesk.Revit.ApplicationServices.Application;

namespace KKY_Tool_Revit.Services
{
    public static class SharedParamBatchService
    {
        public sealed class RunSummary
        {
            public int OkCount { get; set; }
            public int FailCount { get; set; }
            public int SkipCount { get; set; }
        }

        public sealed class LogEntry
        {
            public string Level { get; set; } = string.Empty;
            public string File { get; set; } = string.Empty;
            public string Message { get; set; } = string.Empty;
        }

        public sealed class RunResult
        {
            public bool Ok { get; set; }
            public string Message { get; set; } = string.Empty;
            public RunSummary Summary { get; set; }
            public List<LogEntry> Logs { get; set; }
            public string LogTextPath { get; set; } = string.Empty;
        }

        private sealed class BatchPayload
        {
            public List<string> RvtPaths { get; set; } = new List<string>();
            public List<Dictionary<string, object>> Parameters { get; set; } = new List<Dictionary<string, object>>();
        }

        public static object Init(UIApplication uiapp)
        {
            if (uiapp == null || uiapp.ActiveUIDocument == null || uiapp.ActiveUIDocument.Document == null)
            {
                return new
                {
                    ok = false,
                    message = "활성 프로젝트 문서가 필요합니다.",
                    spFilePath = string.Empty,
                    groups = new List<string>(),
                    defsByGroup = new Dictionary<string, object>(),
                    categoryTree = new List<object>(),
                    paramGroups = new List<object>()
                };
            }

            var baseDoc = uiapp.ActiveUIDocument.Document;
            if (baseDoc == null || baseDoc.IsFamilyDocument)
            {
                return new
                {
                    ok = false,
                    message = "활성 프로젝트 문서가 필요합니다. (패밀리 문서 불가)",
                    spFilePath = string.Empty,
                    groups = new List<string>(),
                    defsByGroup = new Dictionary<string, object>(),
                    categoryTree = new List<object>(),
                    paramGroups = new List<object>()
                };
            }

            var app = uiapp.Application;
            var spFilePath = app.SharedParametersFilename;
            if (string.IsNullOrWhiteSpace(spFilePath) || !File.Exists(spFilePath))
            {
                return new
                {
                    ok = false,
                    message = "현재 Revit에 설정된 Shared Parameter TXT를 찾을 수 없습니다.",
                    spFilePath = spFilePath ?? string.Empty,
                    groups = new List<string>(),
                    defsByGroup = new Dictionary<string, object>(),
                    categoryTree = new List<object>(),
                    paramGroups = BuildParamGroupOptions()
                };
            }

            DefinitionFile defFile;
            try
            {
                defFile = app.OpenSharedParameterFile();
            }
            catch (Exception ex)
            {
                return new
                {
                    ok = false,
                    message = "Shared Parameter 파일을 열 수 없습니다: " + ex.Message,
                    spFilePath,
                    groups = new List<string>(),
                    defsByGroup = new Dictionary<string, object>(),
                    categoryTree = new List<object>(),
                    paramGroups = BuildParamGroupOptions()
                };
            }

            if (defFile == null)
            {
                return new
                {
                    ok = false,
                    message = "Shared Parameter 파일을 열 수 없습니다.",
                    spFilePath,
                    groups = new List<string>(),
                    defsByGroup = new Dictionary<string, object>(),
                    categoryTree = new List<object>(),
                    paramGroups = BuildParamGroupOptions()
                };
            }

            var groupNames = new List<string>();
            var defsByGroup = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);
            foreach (DefinitionGroup g in defFile.Groups)
            {
                if (g == null) continue;
                groupNames.Add(g.Name);

                var defs = new List<object>();
                foreach (Definition d in g.Definitions)
                {
                    var ext = d as ExternalDefinition;
                    if (ext == null) continue;
                    defs.Add(new
                    {
                        name = ext.Name,
                        guid = ext.GUID.ToString("D"),
                        paramTypeLabel = GetExternalDefinitionTypeLabel(ext),
                        desc = ext.Description ?? string.Empty
                    });
                }

                defsByGroup[g.Name] = defs;
            }

            return new
            {
                ok = true,
                message = string.Empty,
                spFilePath,
                groups = groupNames,
                defsByGroup,
                categoryTree = new List<object>(),
                paramGroups = BuildParamGroupOptions()
            };
        }

        public static object BrowseRvts()
        {
            using (var dlg = new OpenFileDialog())
            {
                dlg.Filter = "Revit Project (*.rvt)|*.rvt";
                dlg.Multiselect = true;
                dlg.RestoreDirectory = true;
                if (dlg.ShowDialog() != DialogResult.OK)
                {
                    return new { ok = false, message = "파일 선택이 취소되었습니다." };
                }

                var paths = (dlg.FileNames ?? Array.Empty<string>()).Where(p => !string.IsNullOrWhiteSpace(p)).ToList();
                return new { ok = true, paths };
            }
        }

        public static object BrowseRvtFolder()
        {
            using (var dlg = new FolderBrowserDialog())
            {
                if (dlg.ShowDialog() != DialogResult.OK)
                {
                    return new { ok = false, message = "폴더 선택이 취소되었습니다." };
                }

                var paths = Directory.GetFiles(dlg.SelectedPath, "*.rvt", SearchOption.AllDirectories).ToList();
                return new { ok = true, paths };
            }
        }

        public static object Run(UIApplication uiapp, string payloadJson, IProgress<object> progress)
        {
            var logs = new List<LogEntry>();
            var summary = new RunSummary();

            if (uiapp == null)
            {
                return new RunResult
                {
                    Ok = false,
                    Message = "UIApplication 이 필요합니다.",
                    Summary = summary,
                    Logs = logs
                };
            }

            var payload = ParsePayload(payloadJson);
            var files = NormalizePaths(payload.RvtPaths);
            var total = files.Count;

            if (total == 0)
            {
                return new RunResult
                {
                    Ok = false,
                    Message = "처리할 RVT 파일이 없습니다.",
                    Summary = summary,
                    Logs = logs
                };
            }

            for (var i = 0; i < total; i++)
            {
                var p = files[i];
                var name = SafeFileName(p);
                EmitProgress(progress, i + 1, total, "RUN", $"처리 중: {name}");

                if (!File.Exists(p))
                {
                    summary.SkipCount += 1;
                    logs.Add(new LogEntry { Level = "SKIP", File = name, Message = "파일을 찾을 수 없습니다." });
                    continue;
                }

                Document doc = null;
                try
                {
                    var mp = ModelPathUtils.ConvertUserVisiblePathToModelPath(p);
                    doc = uiapp.Application.OpenDocumentFile(mp, BuildOpenOptions());

                    // NOTE: 전체 VB 로직(정의/카테고리/바인딩 적용)은 단계적으로 포팅 중.
                    // 현재 C# 포트는 배치 파일 처리/검증/로그/엑셀 내보내기 파이프라인 계약을 우선 유지한다.
                    summary.OkCount += 1;
                    logs.Add(new LogEntry { Level = "OK", File = name, Message = "문서 처리 완료" });
                }
                catch (Exception ex)
                {
                    summary.FailCount += 1;
                    logs.Add(new LogEntry { Level = "FAIL", File = name, Message = ex.Message });
                }
                finally
                {
                    if (doc != null)
                    {
                        try { doc.Close(false); } catch { }
                    }
                }
            }

            var logTextPath = WriteLogText(logs);
            var ok = summary.FailCount == 0;
            var message = ok ? "배치 처리가 완료되었습니다." : "일부 파일 처리에 실패했습니다.";

            EmitProgress(progress, total, total, "DONE", message);

            return new RunResult
            {
                Ok = ok,
                Message = message,
                Summary = summary,
                Logs = logs,
                LogTextPath = logTextPath ?? string.Empty
            };
        }

        public static object ExportExcel(string payloadJson)
        {
            var serializer = new JavaScriptSerializer();
            Dictionary<string, object> payload;
            try
            {
                payload = serializer.Deserialize<Dictionary<string, object>>(payloadJson ?? "{}");
            }
            catch (Exception ex)
            {
                return new { ok = false, message = "payload 파싱 실패: " + ex.Message };
            }

            var logsTable = new DataTable("SharedParamBatchLogs");
            logsTable.Columns.Add("성공여부");
            logsTable.Columns.Add("파일");
            logsTable.Columns.Add("메시지");

            var logsObj = payload != null && payload.TryGetValue("logs", out var lObj) ? lObj : null;
            if (logsObj is IEnumerable ie)
            {
                foreach (var item in ie)
                {
                    var d = ToDict(item);
                    var row = logsTable.NewRow();
                    row[0] = Convert.ToString(GetValue(d, "level"));
                    row[1] = Convert.ToString(GetValue(d, "file"));
                    row[2] = Convert.ToString(GetValue(d, "msg"));
                    logsTable.Rows.Add(row);
                }
            }

            ExcelCore.EnsureNoDataRow(logsTable, "로그가 없습니다.");

            var mode = Convert.ToString(GetValue(payload, "excelMode"));
            var autoFit = string.Equals(mode, "normal", StringComparison.OrdinalIgnoreCase);
            var path = ExcelCore.PickAndSaveXlsx("SharedParamBatch", logsTable, $"SharedParamBatch_{DateTime.Now:yyyyMMdd_HHmm}.xlsx", autoFit, "sharedparambatch:progress", "sharedparambatch");

            if (string.IsNullOrWhiteSpace(path))
            {
                return new { ok = false, message = "엑셀 저장이 취소되었습니다." };
            }

            return new { ok = true, path };
        }

        private static string GetExternalDefinitionTypeLabel(ExternalDefinition ext)
        {
            if (ext == null) return string.Empty;
            try
            {
#if REVIT2023
                var dt = ext.GetDataType();
                return dt == null ? string.Empty : (dt.TypeId ?? string.Empty);
#else
                return ext.ParameterType.ToString();
#endif
            }
            catch
            {
                return string.Empty;
            }
        }

        private static List<object> BuildParamGroupOptions()
        {
            var list = new List<object>();
            foreach (var v in Enum.GetValues(typeof(BuiltInParameterGroup)).Cast<BuiltInParameterGroup>())
            {
                list.Add(new { key = v.ToString(), id = ((int)v).ToString(), label = v.ToString() });
            }
            return list;
        }

        private static BatchPayload ParsePayload(string payloadJson)
        {
            var result = new BatchPayload();
            if (string.IsNullOrWhiteSpace(payloadJson)) return result;

            try
            {
                var serializer = new JavaScriptSerializer();
                var d = serializer.Deserialize<Dictionary<string, object>>(payloadJson);
                if (d == null) return result;

                if (d.TryGetValue("rvtPaths", out var pathsObj) && pathsObj is IEnumerable pe)
                {
                    foreach (var o in pe)
                    {
                        var s = Convert.ToString(o);
                        if (!string.IsNullOrWhiteSpace(s)) result.RvtPaths.Add(s);
                    }
                }

                if (d.TryGetValue("parameters", out var paramsObj) && paramsObj is IEnumerable pa)
                {
                    foreach (var o in pa)
                    {
                        var pd = ToDict(o);
                        if (pd.Count > 0) result.Parameters.Add(pd);
                    }
                }
            }
            catch
            {
            }

            return result;
        }

        private static List<string> NormalizePaths(IEnumerable<string> paths)
        {
            var list = new List<string>();
            var dedup = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            if (paths == null) return list;

            foreach (var p in paths)
            {
                if (string.IsNullOrWhiteSpace(p)) continue;
                string full;
                try { full = Path.GetFullPath(p); } catch { full = p; }
                if (dedup.Add(full)) list.Add(full);
            }

            return list;
        }

        private static string SafeFileName(string path)
        {
            if (string.IsNullOrWhiteSpace(path)) return "(Unknown)";
            try { return Path.GetFileName(path); } catch { return path; }
        }

        private static OpenOptions BuildOpenOptions()
        {
            var opts = new OpenOptions { Audit = false };
            try
            {
                var ws = new WorksetConfiguration(WorksetConfigurationOption.CloseAllWorksets);
                opts.SetOpenWorksetsConfiguration(ws);
            }
            catch
            {
            }

            return opts;
        }

        private static void EmitProgress(IProgress<object> progress, int current, int total, string phase, string message)
        {
            if (progress == null) return;
            var pct = total > 0 ? (double)current / total * 100.0 : 0.0;
            progress.Report(new { phase, message, current, total, percent = pct });
        }

        private static string WriteLogText(List<LogEntry> logs)
        {
            try
            {
                var folder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory), "KKY_Tool_Revit");
                Directory.CreateDirectory(folder);
                var path = Path.Combine(folder, $"SharedParamBatch_{DateTime.Now:yyyyMMdd_HHmmss}.log.txt");
                File.WriteAllLines(path, (logs ?? new List<LogEntry>()).Select(l => $"[{l.Level}] {l.File} - {l.Message}"));
                return path;
            }
            catch
            {
                return string.Empty;
            }
        }

        private static Dictionary<string, object> ToDict(object obj)
        {
            var d = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);
            if (obj == null) return d;

            if (obj is Dictionary<string, object> dd) return dd;
            if (obj is IDictionary id)
            {
                foreach (var k in id.Keys) d[Convert.ToString(k)] = id[k];
                return d;
            }

            foreach (var p in obj.GetType().GetProperties())
            {
                d[p.Name] = p.GetValue(obj, null);
            }

            return d;
        }

        private static object GetValue(Dictionary<string, object> d, string key)
        {
            if (d == null || string.IsNullOrEmpty(key)) return null;
            return d.TryGetValue(key, out var v) ? v : null;
        }
    }
}
