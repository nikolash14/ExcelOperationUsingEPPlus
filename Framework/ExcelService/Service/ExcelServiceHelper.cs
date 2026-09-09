using ExcelService.Model;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
namespace ExcelService.Service
{
    public class ExcelServiceHelper
    {
        // Returns the maximum depth (number of header rows) required for the provided headers
        public static int GetTotalDepth(List<ExcelHeader> headers)
        {
            if (headers == null || headers.Count == 0)
                return 0;
            return headers.Max(h => h.GetDepth());
        }

        public static byte[] Export(
            List<ExcelHeader> headers,
            List<List<object>> data)
        {
            int totalDepth = GetTotalDepth(headers);
            ExcelPackage.License.SetNonCommercialOrganization("YourOrganizationName");
            using (var package = new ExcelPackage())
            {
                var ws = package.Workbook.Worksheets.Add("Sheet1");
                // Render headers
                int col = 1;
                foreach (var header in headers)
                {
                    RenderHeaders(ws, header, 1, col, totalDepth);
                    col += header.GetLeafCount();
                }

                // Render data rows
                int row = headers.Max(h => h.GetDepth()) + 1;
                foreach (var record in data)
                {
                    for (int i = 0; i < record.Count; i++)
                    {
                        ws.Cells[row, i + 1].Value = record[i];
                    }
                    row++;
                }

                ws.Cells[ws.Dimension.Address].AutoFitColumns();
                // ... inside ExcelServiceHelper.Export, after the data rows loop
                int dataStartRow = headers.Max(h => h.GetDepth()) + 1;
                int lastRow = row - 1; // row was incremented after last record
                int lastCol = data.Count > 0 ? data[0].Count : headers.Sum(h => h.GetLeafCount());

                // Apply thin borders to every cell in the data area
                var dataRange = ws.Cells[1, 1, lastRow, lastCol];
                dataRange.Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                dataRange.Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                dataRange.Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                dataRange.Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                return package.GetAsByteArray();
            }
        }

        private static void RenderHeaders(
            ExcelWorksheet ws,
            ExcelHeader header,
            int row,
            int col,
            int totalDepth,
            int usedDepth = 0)
        {
            int leafCount = header.GetLeafCount();
            int depth = header.GetDepth();

            // If header has subheaders, only span one row vertically; otherwise span full depth
            bool hasChildren = header.SubHeaders != null && header.SubHeaders.Any();
            int rowSpan = hasChildren ? 1 : totalDepth - usedDepth;
            int colSpan = leafCount;

            var range = ws.Cells[row, col, row + rowSpan - 1, col + colSpan - 1];

            // Only merge if the range covers multiple cells
            if (rowSpan > 1 || colSpan > 1)
            {
                range.Merge = true;
            }
            range.Value = header.Name;

            // Apply style
            range.Style.Font.Bold = header.Style.Bold;
            range.Style.Font.Size = header.Style.FontSize;
            range.Style.Font.Color.SetColor(header.Style.FontColor);
            range.Style.Fill.PatternType = ExcelFillStyle.Solid;
            range.Style.Fill.BackgroundColor.SetColor(header.Style.BackgroundColor);
            range.Style.HorizontalAlignment = header.Style.HorizontalAlignment;
            range.Style.VerticalAlignment = header.Style.VerticalAlignment;

            usedDepth++;
            // Render children
            int childCol = col;
            foreach (var child in header.SubHeaders)
            {
                RenderHeaders(ws, child, row + 1, childCol, totalDepth, usedDepth);
                childCol += child.GetLeafCount();
            }
        }
    }
}
