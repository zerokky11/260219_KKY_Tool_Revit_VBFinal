using System;
using System.Runtime.CompilerServices;
using NPOI.SS.UserModel;

namespace KKY_Tool_Revit.Infrastructure
{
    public static class ExcelStyleHelper
    {
        public enum RowStatus
        {
            None = 0,
            Info = 1,
            Warning = 2,
            Error = 3
        }

        private sealed class StyleSet
        {
            public ICellStyle HeaderStyle;
            public ICellStyle InfoStyle;
            public ICellStyle WarningStyle;
            public ICellStyle ErrorStyle;
        }

        private static readonly ConditionalWeakTable<IWorkbook, StyleSet> WbStyles = new ConditionalWeakTable<IWorkbook, StyleSet>();
        private static readonly ConditionalWeakTable<IWorkbook, StyleSet> WbNoWrapStyles = new ConditionalWeakTable<IWorkbook, StyleSet>();

        private static StyleSet GetStyleSet(IWorkbook wb)
        {
            if (wb == null) throw new ArgumentNullException(nameof(wb));
            return WbStyles.GetValue(wb, CreateStyleSet);
        }

        private static StyleSet GetNoWrapStyleSet(IWorkbook wb)
        {
            if (wb == null) throw new ArgumentNullException(nameof(wb));
            return WbNoWrapStyles.GetValue(wb, CreateStyleSet);
        }

        private static StyleSet CreateStyleSet(IWorkbook wb)
        {
            var setx = new StyleSet();

            var header = wb.CreateCellStyle();
            var hf = wb.CreateFont();
            hf.IsBold = true;
            hf.Color = IndexedColors.White.Index;
            header.SetFont(hf);
            header.Alignment = HorizontalAlignment.Center;
            header.VerticalAlignment = VerticalAlignment.Center;
            header.WrapText = false;
            header.FillForegroundColor = IndexedColors.Grey50Percent.Index;
            header.FillPattern = FillPattern.SolidForeground;
            header.BorderBottom = BorderStyle.Thin;
            header.BorderTop = BorderStyle.Thin;
            header.BorderLeft = BorderStyle.Thin;
            header.BorderRight = BorderStyle.Thin;
            setx.HeaderStyle = header;

            setx.InfoStyle = CreateFillStyle(wb, IndexedColors.LightCornflowerBlue.Index);
            setx.WarningStyle = CreateFillStyle(wb, IndexedColors.LightYellow.Index);
            setx.ErrorStyle = CreateFillStyle(wb, IndexedColors.Rose.Index);

            return setx;
        }

        private static ICellStyle CreateFillStyle(IWorkbook wb, short fillColor)
        {
            var s = wb.CreateCellStyle();
            s.WrapText = false;
            s.VerticalAlignment = VerticalAlignment.Center;
            s.FillForegroundColor = fillColor;
            s.FillPattern = FillPattern.SolidForeground;
            s.BorderBottom = BorderStyle.Thin;
            s.BorderTop = BorderStyle.Thin;
            s.BorderLeft = BorderStyle.Thin;
            s.BorderRight = BorderStyle.Thin;
            return s;
        }

        public static ICellStyle GetHeaderStyle(IWorkbook wb) => GetStyleSet(wb).HeaderStyle;

        public static ICellStyle GetHeaderStyleNoWrap(IWorkbook wb) => GetNoWrapStyleSet(wb).HeaderStyle;

        public static ICellStyle GetRowStyle(IWorkbook wb, RowStatus status)
        {
            var setx = GetStyleSet(wb);
            switch (status)
            {
                case RowStatus.Info: return setx.InfoStyle;
                case RowStatus.Warning: return setx.WarningStyle;
                case RowStatus.Error: return setx.ErrorStyle;
                default: return null;
            }
        }

        public static ICellStyle GetRowStyleNoWrap(IWorkbook wb, RowStatus status)
        {
            var setx = GetNoWrapStyleSet(wb);
            switch (status)
            {
                case RowStatus.Info: return setx.InfoStyle;
                case RowStatus.Warning: return setx.WarningStyle;
                case RowStatus.Error: return setx.ErrorStyle;
                default: return null;
            }
        }

        public static void ApplyStyleToRow(IRow row, int colCount, ICellStyle style)
        {
            if (row == null || style == null) return;

            var lastCol = Math.Max(0, colCount - 1);
            for (var c = 0; c <= lastCol; c++)
            {
                var cell = row.GetCell(c) ?? row.CreateCell(c);
                if (cell.CellType == CellType.Blank)
                {
                    cell.SetCellValue(string.Empty);
                }
                cell.CellStyle = style;
            }
        }
    }
}
