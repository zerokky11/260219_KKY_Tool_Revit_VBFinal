using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using KKY_Tool_Revit.Infrastructure;
using KKY_Tool_Revit.Services;

namespace KKY_Tool_Revit.Exports
{
    public static class FamilyLinkAuditExport
    {
        public static readonly string[] Schema =
        {
            "FileName",
            "HostFamilyName",
            "HostFamilyCategory",
            "NestedFamilyName",
            "NestedTypeName",
            "NestedCategory",
            "NestedParamName",
            "TargetParamName",
            "ExpectedGuid",
            "FoundScope",
            "NestedParamGuid",
            "NestedParamDataType",
            "AssocHostParamName",
            "HostParamIsShared",
            "Issue",
            "Notes"
        };

        public static string Export(IEnumerable<FamilyLinkAuditRow> rows, bool fastExport = true, bool autoFit = false)
        {
            if (rows == null) return string.Empty;
            var table = ToDataTable(rows);
            global::KKY_Tool_Revit.Infrastructure.ResultTableFilter.KeepOnlyIssues("familylink", table);
            ExcelCore.EnsureMessageRow(table, "오류가 없습니다.");
            if (!ValidateSchema(table))
            {
                throw new InvalidOperationException("스키마 검증 실패: 컬럼 순서/헤더가 규격과 다릅니다.");
            }

            var defaultName = $"FamilyLinkAudit_{DateTime.Now:yyyyMMdd_HHmm}.xlsx";
            using (var dlg = new SaveFileDialog())
            {
                dlg.Filter = "Excel Workbook (*.xlsx)|*.xlsx|CSV (*.csv)|*.csv";
                dlg.FileName = defaultName;
                dlg.AddExtension = true;
                dlg.DefaultExt = "xlsx";
                dlg.OverwritePrompt = true;
                dlg.RestoreDirectory = true;

                if (dlg.ShowDialog() != DialogResult.OK) return string.Empty;

                var filePath = dlg.FileName;
                var ext = Path.GetExtension(filePath).ToLowerInvariant();
                if (string.IsNullOrWhiteSpace(ext))
                {
                    ext = dlg.FilterIndex == 2 ? ".csv" : ".xlsx";
                    filePath += ext;
                }

                if (ext == ".csv")
                {
                    SaveCsv(filePath, table);
                }
                else
                {
                    var doAutoFit = !fastExport && autoFit;
                    var excelMode = fastExport ? "fast" : "normal";
                    ExcelCore.SaveXlsx(filePath, "FamilyLinkAudit", table, doAutoFit);
                    ExcelExportStyleRegistry.ApplyStylesForKey("familylink", filePath, autoFit: doAutoFit, excelMode: excelMode);
                }

                return filePath;
            }
        }

        public static DataTable ToDataTable(IEnumerable<FamilyLinkAuditRow> rows)
        {
            return BuildTable(rows);
        }

        private static DataTable BuildTable(IEnumerable<FamilyLinkAuditRow> rows)
        {
            var dt = new DataTable("FamilyLinkAudit");
            foreach (var h in Schema)
            {
                dt.Columns.Add(h);
            }

            foreach (var r in rows)
            {
                var dr = dt.NewRow();
                dr["FileName"] = SafeStr(r.FileName);
                dr["HostFamilyName"] = SafeStr(r.HostFamilyName);
                dr["HostFamilyCategory"] = SafeStr(r.HostFamilyCategory);
                dr["NestedFamilyName"] = SafeStr(r.NestedFamilyName);
                dr["NestedTypeName"] = SafeStr(r.NestedTypeName);
                dr["NestedCategory"] = SafeStr(r.NestedCategory);
                dr["NestedParamName"] = SafeStr(r.NestedParamName);
                dr["TargetParamName"] = SafeStr(r.TargetParamName);
                dr["ExpectedGuid"] = SafeStr(r.ExpectedGuid);
                dr["FoundScope"] = SafeStr(r.FoundScope);
                dr["NestedParamGuid"] = SafeStr(r.NestedParamGuid);
                dr["NestedParamDataType"] = SafeStr(r.NestedParamDataType);
                dr["AssocHostParamName"] = SafeStr(r.AssocHostParamName);
                dr["HostParamIsShared"] = SafeStr(r.HostParamIsShared);
                dr["Issue"] = SafeStr(r.Issue);
                dr["Notes"] = SafeStr(r.Notes);
                dt.Rows.Add(dr);
            }

            return dt;
        }

        private static bool ValidateSchema(DataTable table)
        {
            if (table == null) return false;
            if (table.Columns.Count != Schema.Length) return false;

            for (var i = 0; i < Schema.Length; i++)
            {
                if (!string.Equals(table.Columns[i].ColumnName, Schema[i], StringComparison.Ordinal))
                {
                    return false;
                }
            }

            return true;
        }

        private static void SaveCsv(string filePath, DataTable table)
        {
            using (var fs = new FileStream(filePath, FileMode.Create, FileAccess.Write, FileShare.Read))
            using (var sw = new StreamWriter(fs, new UTF8Encoding(true)))
            {
                sw.WriteLine(string.Join(",", Schema.Select(CsvEscape)));
                foreach (DataRow row in table.Rows)
                {
                    var cols = new List<string>();
                    foreach (var h in Schema)
                    {
                        cols.Add(CsvEscape((row[h] ?? string.Empty).ToString()));
                    }

                    sw.WriteLine(string.Join(",", cols));
                }
            }
        }

        private static string CsvEscape(string s)
        {
            if (s == null) s = string.Empty;
            const string quote = "\"";
            var needsQuotes = s.Contains(',') || s.Contains(quote) || s.Contains("\r") || s.Contains("\n");
            s = s.Replace(quote, quote + quote);
            return needsQuotes ? quote + s + quote : s;
        }

        private static string SafeStr(string s)
        {
            return s ?? string.Empty;
        }
    }
}
