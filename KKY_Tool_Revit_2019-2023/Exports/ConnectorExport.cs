using System;
using System.Data;

namespace KKY_Tool_Revit.Exports
{
    public static class ConnectorExport
    {
        public static string SaveWithDialog(DataTable resultTable)
        {
            if (resultTable == null) return string.Empty;
            EnsureMessageRow(resultTable);

            var outPath = global::KKY_Tool_Revit.Infrastructure.ExcelCore.PickAndSaveXlsx(
                "Connector Diagnostics",
                resultTable,
                "ConnectorDiagnostics.xlsx");

            if (string.IsNullOrWhiteSpace(outPath)) return outPath;

            try
            {
                global::KKY_Tool_Revit.Infrastructure.ExcelExportStyleRegistry.ApplyStylesForKey("connector", outPath);
            }
            catch
            {
                // ignore
            }

            return outPath;
        }

        public static void Save(string outPath, DataTable resultTable)
        {
            if (string.IsNullOrWhiteSpace(outPath)) return;
            if (resultTable == null) return;
            EnsureMessageRow(resultTable);

            global::KKY_Tool_Revit.Infrastructure.ExcelCore.SaveXlsx(outPath, "Connector Diagnostics", resultTable);

            try
            {
                global::KKY_Tool_Revit.Infrastructure.ExcelExportStyleRegistry.ApplyStylesForKey("connector", outPath);
            }
            catch
            {
                // ignore
            }
        }

        private static void EnsureMessageRow(DataTable table)
        {
            global::KKY_Tool_Revit.Infrastructure.ExcelCore.EnsureMessageRow(table, "오류가 없습니다.");
        }
    }
}
