using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using Autodesk.Revit.UI;
using KKY_Tool_Revit.Services;

namespace KKY_Tool_Revit.UI.Hub
{
    public sealed partial class UiBridgeExternalEvent
    {
        private List<SegmentPmsCheckService.PmsRow> _pmsRows;
        private DataSet _extractData;
        private SegmentPmsCheckService.RunResult _lastRunResult;

        private static SegmentPmsCheckService.ExtractOptions ParseExtractOptions(object payload)
        {
            return new SegmentPmsCheckService.ExtractOptions
            {
                IncludeLinks = GetBool(payload, "includeLinks", false),
                IncludeRouting = GetBool(payload, "includeRouting", true)
            };
        }

        private static SegmentPmsCheckService.CompareOptions ParseCompareOptions(object payload)
        {
            return new SegmentPmsCheckService.CompareOptions
            {
                StrictSize = GetBool(payload, "strictSize", false),
                StrictRouting = GetBool(payload, "strictRouting", false)
            };
        }

        private static List<SegmentPmsCheckService.MappingSelection> ParseMappings(object payload)
        {
            var res = new List<SegmentPmsCheckService.MappingSelection>();
            var raw = GetProp(payload, "mappings");
            if (raw is IEnumerable<object> arr)
            {
                foreach (var it in arr)
                {
                    res.Add(new SegmentPmsCheckService.MappingSelection
                    {
                        GroupKey = GetString(it, "group", ""),
                        Class = GetString(it, "cls", ""),
                        SegmentKey = GetString(it, "segment", "")
                    });
                }
            }
            return res;
        }

        private static void ReportProgress(double pct, string text)
        {
            SendToWeb("segmentpms:progress", new { pct, text });
        }

        private void HandleSegmentPmsRvtPickFiles(UIApplication app, object payload)
        {
            var files = PickFiles("RVT (*.rvt)|*.rvt", true);
            SendToWeb("segmentpms:rvt-picked-files", new { paths = files });
        }

        private void HandleSegmentPmsRvtPickFolder(UIApplication app, object payload)
        {
            var folder = PickFolder();
            var files = string.IsNullOrWhiteSpace(folder) ? new List<string>() : Directory.GetFiles(folder, "*.rvt", SearchOption.TopDirectoryOnly).ToList();
            SendToWeb("segmentpms:rvt-picked-folder", new { paths = files });
        }

        private void HandleSegmentPmsExtractStart(UIApplication app, object payload)
        {
            var files = GetStringList(payload, "files");
            if (files.Count == 0)
            {
                SendToWeb("segmentpms:error", new { message = "추출할 RVT 파일을 선택하세요." });
                return;
            }

            var opts = ParseExtractOptions(payload);
            _extractData = SegmentPmsCheckService.ExtractToDataSet(app, files, opts, ReportProgress);
            var groups = SegmentPmsCheckService.BuildGroups(_extractData);
            var suggest = SegmentPmsCheckService.SuggestGroupMappings(groups, _pmsRows ?? new List<SegmentPmsCheckService.PmsRow>());
            SendToWeb("segmentpms:extract-done", new { summary = new { files = files.Count, groups = groups.Count }, groups, suggestions = suggest });
        }

        private void HandleSegmentPmsSaveExtract(UIApplication app, object payload)
        {
            if (_extractData == null)
            {
                SendToWeb("segmentpms:error", new { message = "저장할 추출 데이터가 없습니다." });
                return;
            }

            var path = PickSaveXlsx("segment_extract.xlsx");
            if (string.IsNullOrWhiteSpace(path)) return;
            SegmentPmsCheckService.SaveDataSetToXlsx(_extractData, path, ParseExcelMode(payload), "segmentpms:progress");
            SendToWeb("segmentpms:extract-saved", new { path });
        }

        private void HandleSegmentPmsLoadExtract(UIApplication app, object payload)
        {
            var path = GetString(payload, "path", string.Empty);
            if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
            {
                path = PickFile("Excel (*.xlsx)|*.xlsx");
            }
            if (string.IsNullOrWhiteSpace(path)) return;

            _extractData = SegmentPmsCheckService.LoadExtractFromXlsx(path);
            var groups = SegmentPmsCheckService.BuildGroups(_extractData);
            var suggest = SegmentPmsCheckService.SuggestGroupMappings(groups, _pmsRows ?? new List<SegmentPmsCheckService.PmsRow>());
            SendToWeb("segmentpms:extract-loaded", new { path, groups, suggestions = suggest });
        }

        private void HandleSegmentPmsRegisterPms(UIApplication app, object payload)
        {
            var path = GetString(payload, "path", string.Empty);
            if (string.IsNullOrWhiteSpace(path) || !File.Exists(path)) path = PickFile("Excel (*.xlsx)|*.xlsx");
            if (string.IsNullOrWhiteSpace(path)) { SendToWeb("segmentpms:error", new { message = "PMS 불러오기가 취소되었습니다." }); return; }
            _pmsRows = SegmentPmsCheckService.LoadPmsRowsFromXlsx(path);
            SendToWeb("segmentpms:pms-registered", new { path, count = _pmsRows.Count });
        }

        private void HandleSegmentPmsExportTemplate(UIApplication app, object payload)
        {
            var p = PickSaveXlsx("segment_pms_template.xlsx");
            if (string.IsNullOrWhiteSpace(p)) return;
            var ds = new DataSet();
            var t = new DataTable("PMS_TEMPLATE");
            t.Columns.Add("CLASS"); t.Columns.Add("SEGMENT"); t.Columns.Add("SIZE"); t.Columns.Add("ROUTING");
            ds.Tables.Add(t);
            SegmentPmsCheckService.SaveDataSetToXlsx(ds, p, ParseExcelMode(payload), "segmentpms:progress");
            SendToWeb("segmentpms:pms-template-exported", new { path = p });
        }

        private void HandleSegmentPmsPrepareMapping(UIApplication app, object payload)
        {
            var groups = SegmentPmsCheckService.BuildGroups(_extractData);
            var suggest = SegmentPmsCheckService.SuggestGroupMappings(groups, _pmsRows ?? new List<SegmentPmsCheckService.PmsRow>());
            SendToWeb("segmentpms:mapping-prepared", new { groups, suggestions = suggest });
        }

        private void HandleSegmentPmsRun(UIApplication app, object payload)
        {
            if (_extractData == null) { SendToWeb("segmentpms:error", new { message = "추출 데이터가 없습니다." }); return; }
            var mappings = ParseMappings(payload);
            var compare = ParseCompareOptions(payload);
            _lastRunResult = SegmentPmsCheckService.RunCompare(_extractData, _pmsRows ?? new List<SegmentPmsCheckService.PmsRow>(), mappings, compare);
            SendToWeb("segmentpms:done", new { summary = _lastRunResult?.Summary, mapRows = _lastRunResult?.MapTable?.Rows.Count ?? 0, compareRows = _lastRunResult?.CompareTable?.Rows.Count ?? 0 });
        }

        private void HandleSegmentPmsSaveResult(UIApplication app, object payload)
        {
            if (_lastRunResult == null) { SendToWeb("segmentpms:error", new { message = "저장할 비교 결과가 없습니다." }); return; }
            var path = PickSaveXlsx("segment_result.xlsx");
            if (string.IsNullOrWhiteSpace(path)) return;
            var dt = _lastRunResult.CompareTable ?? new DataTable("Compare");
            Infrastructure.ExcelCore.SaveXlsx(path, dt.TableName, dt, ParseExcelMode(payload), sheetKey: dt.TableName, progressKey: "segmentpms:progress", exportKind: "segmentpms");
            SendToWeb("segmentpms:result-saved", new { path });
        }

        private static List<string> PickFiles(string filter, bool multiselect)
        {
            var dlg = new Microsoft.Win32.OpenFileDialog { Filter = filter, Multiselect = multiselect };
            return dlg.ShowDialog() == true ? dlg.FileNames.ToList() : new List<string>();
        }

        private static string PickFile(string filter)
        {
            var dlg = new Microsoft.Win32.OpenFileDialog { Filter = filter, Multiselect = false };
            return dlg.ShowDialog() == true ? dlg.FileName : string.Empty;
        }

        private static string PickSaveXlsx(string defaultName)
        {
            var dlg = new Microsoft.Win32.SaveFileDialog { Filter = "Excel (*.xlsx)|*.xlsx", FileName = defaultName };
            return dlg.ShowDialog() == true ? dlg.FileName : string.Empty;
        }

        private static string PickFolder()
        {
            using (var dlg = new System.Windows.Forms.FolderBrowserDialog())
            {
                return dlg.ShowDialog() == System.Windows.Forms.DialogResult.OK ? dlg.SelectedPath : string.Empty;
            }
        }
    }
}
