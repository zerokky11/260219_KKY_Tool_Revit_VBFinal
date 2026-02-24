using System;
using System.Collections.Generic;
using System.Data;

namespace KKY_Tool_Revit.Infrastructure
{
    public static class ResultTableFilter
    {
        public static void KeepOnlyIssues(string key, DataTable table)
        {
            if (table == null) return;
            if (table.Rows == null || table.Rows.Count == 0) return;

            for (var i = table.Rows.Count - 1; i >= 0; i--)
            {
                var st = ExcelExportStyleRegistry.Resolve(key, table.Rows[i], table);
                var keep = st == ExcelStyleHelper.RowStatus.Warning || st == ExcelStyleHelper.RowStatus.Error;
                if (!keep)
                {
                    table.Rows.RemoveAt(i);
                }
            }
        }

        public static void KeepOnlyByNameSet(DataTable table, string nameColumn, HashSet<string> keepNames)
        {
            if (table == null) return;
            if (keepNames == null) return;
            if (table.Rows == null || table.Rows.Count == 0) return;
            if (!table.Columns.Contains(nameColumn)) return;

            for (var i = table.Rows.Count - 1; i >= 0; i--)
            {
                var nm = string.Empty;
                if (table.Rows[i] != null && table.Rows[i][nameColumn] != null)
                {
                    nm = Convert.ToString(table.Rows[i][nameColumn]).Trim();
                }

                if (string.IsNullOrWhiteSpace(nm) || !keepNames.Contains(nm))
                {
                    table.Rows.RemoveAt(i);
                }
            }
        }
    }
}
