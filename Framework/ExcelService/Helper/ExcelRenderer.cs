using ExcelService.Model;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Linq;

namespace ExcelService.Helper
{
    internal static class ExcelRenderer
    {
        internal static void AutoFitColumns(
            ExcelWorkbookStyle excelWorkbookStyle,
            ExcelWorksheet ws)
        {
            if (excelWorkbookStyle != null && excelWorkbookStyle.AutoFitColumns && ws?.Dimension != null)
                ws.Cells[ws.Dimension.Address].AutoFitColumns();
        }

        internal static void ApplyBorders(
            ExcelWorkbookStyle excelWorkbookStyle,
            ExcelWorksheet ws,
            int lastRow,
            int lastCol)
        {
            if (excelWorkbookStyle != null && excelWorkbookStyle.ApplyBorders && ws?.Dimension != null)
            {
                var dataRange = ws.Cells[1, 1, lastRow, lastCol];
                dataRange.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                dataRange.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                dataRange.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                dataRange.Style.Border.Right.Style = ExcelBorderStyle.Thin;
            }
        }

        internal static void FreezePanes(
            ExcelWorkbookStyle excelWorkbookStyle,
            int totalDepth,
            ExcelWorksheet ws)
        {
            if (excelWorkbookStyle != null && excelWorkbookStyle.FreezeHeader)
                ws.View.FreezePanes(totalDepth + 1, excelWorkbookStyle.FreezeColumnCount + 1);
            else
                ws.View.FreezePanes(1, excelWorkbookStyle?.FreezeColumnCount + 1 ?? 1);
        }

        internal static void ApplyExcelProtection(
            ExcelProtectionSetting excelProtectionSetting,
            ExcelPackage package,
            ExcelWorksheet ws,
            int totalDepth)
        {
            if (excelProtectionSetting != null && excelProtectionSetting.IsProtected)
            {
                if (excelProtectionSetting.AutoFilterInCaseProtected && ws?.Dimension != null)
                {
                    ws.Cells[totalDepth, 1, totalDepth, ws.Dimension.End.Column].AutoFilter =
                        excelProtectionSetting.AutoFilterInCaseProtected;
                }

                if (excelProtectionSetting.AllowEditToAllData && ws?.Dimension != null)
                {
                    ws.Cells[totalDepth + 1, 1, ws.Dimension.End.Row, ws.Dimension.End.Column].Style.Locked = false;
                }
                else if (excelProtectionSetting.EditableColumnIndex != null && excelProtectionSetting.EditableColumnIndex.Count > 0)
                {
                    foreach (var colIndex in excelProtectionSetting.EditableColumnIndex)
                    {
                        ws.Cells[totalDepth + 1, colIndex, ws.Dimension.End.Row, colIndex].Style.Locked = false;
                    }
                }

                ws.Protection.IsProtected = true;
                ws.Protection.SetPassword(excelProtectionSetting.Password);

                ws.Protection.AllowDeleteColumns = false;
                ws.Protection.AllowDeleteRows = false;
                ws.Protection.AllowInsertColumns = false;
                ws.Protection.AllowInsertRows = false;

                ws.Protection.AllowSelectLockedCells = true;
                ws.Protection.AllowSelectUnlockedCells = true;
                ws.Protection.AllowFormatCells = true;
                ws.Protection.AllowFormatColumns = true;
                ws.Protection.AllowFormatRows = true;
                ws.Protection.AllowSort = true;
                ws.Protection.AllowAutoFilter = true;
                ws.Protection.AllowPivotTables = true;
            }
        }

        internal static void RenderHeaders(
            ExcelWorksheet ws,
            ExcelHeader header,
            int row,
            int col,
            int totalDepth,
            int usedDepth = 0)
        {
            int leafCount = header.GetLeafCount();
            int depth = header.GetDepth();

            bool hasChildren = header.SubHeaders != null && header.SubHeaders.Any();
            int rowSpan = hasChildren ? 1 : totalDepth - usedDepth;
            int colSpan = leafCount;
            header.CollNo = col;
            var range = ws.Cells[row, col, row + rowSpan - 1, col + colSpan - 1];

            if (rowSpan > 1 || colSpan > 1)
            {
                range.Merge = true;
            }
            range.Value = header.Name;

            range.Style.Font.Bold = header.Style.Bold;
            range.Style.Font.Size = header.Style.FontSize;
            range.Style.Font.Color.SetColor(header.Style.FontColor);
            range.Style.Fill.PatternType = ExcelFillStyle.Solid;
            range.Style.Fill.BackgroundColor.SetColor(header.Style.BackgroundColor);
            range.Style.HorizontalAlignment = header.Style.HorizontalAlignment;
            range.Style.VerticalAlignment = header.Style.VerticalAlignment;

            usedDepth++;
            int childCol = col;
            foreach (var child in header.SubHeaders)
            {
                RenderHeaders(ws, child, row + 1, childCol, totalDepth, usedDepth);
                childCol += child.GetLeafCount();
            }
        }
    }
}
