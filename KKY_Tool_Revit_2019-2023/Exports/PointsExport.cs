using System.Data;
using KKY_Tool_Revit.Infrastructure;

namespace KKY_Tool_Revit.Exports
{
    public static class PointsExport
    {
        public static string SaveWithDialog(DataTable resultTable)
        {
            if (resultTable == null) return string.Empty;
            if (resultTable.Rows.Count == 0)
            {
                ExcelCore.EnsureNoDataRow(resultTable, "추출 결과가 없습니다.");
            }
            return ExcelCore.PickAndSaveXlsx("Exported Points", resultTable, "ExportPoints.xlsx", exportKind: "points");
        }

        public static void Save(string outPath, DataTable resultTable)
        {
            if (resultTable == null) return;
            if (resultTable.Rows.Count == 0)
            {
                ExcelCore.EnsureNoDataRow(resultTable, "추출 결과가 없습니다.");
            }

            ExcelCore.SaveXlsx(outPath, "Exported Points", resultTable, exportKind: "points");
        }
    }
}
