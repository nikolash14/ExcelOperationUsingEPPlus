using ExcelService.Model;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Linq;

namespace ExcelService.Helper
{
    internal class ExcelHelper
    {
        internal static void AutoFitColumns(
            ExcelWorkbookStyle excelWorkbookStyle,
            ExcelWorksheet ws)
        {
            if (excelWorkbookStyle.AutoFitColumns)
                ws.Cells[ws.Dimension.Address].AutoFitColumns();
        }

        internal static void ApplyBorders(
            ExcelWorkbookStyle excelWorkbookStyle,
            ExcelWorksheet ws,
            int lastRow,
            int lastCol)
        {
            if (excelWorkbookStyle.ApplyBorders)
            {
                var dataRange = ws.Cells[1, 1, lastRow, lastCol];
                dataRange.Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                dataRange.Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                dataRange.Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                dataRange.Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            }
        }

        internal static void FreezePanes(
            ExcelWorkbookStyle excelWorkbookStyle,
            int totalDepth,
            ExcelWorksheet ws)
        {
            if (excelWorkbookStyle.FreezeHeader)
                ws.View.FreezePanes(totalDepth + 1, excelWorkbookStyle.FreezeColumnCount + 1);
            else
                ws.View.FreezePanes(1, excelWorkbookStyle.FreezeColumnCount + 1);
        }

        internal static void ApplyExcelProtection(
            ExcelProtectionSetting excelProtectionSetting,
            ExcelPackage package,
            ExcelWorksheet ws,
            int totalDepth)
        {
            if (excelProtectionSetting != null &&
                excelProtectionSetting.IsProtected)
            {
                // Define the autofilter range (e.g., first row as header)
                if (excelProtectionSetting.AutoFilterInCaseProtected)
                {
                    ws.Cells[totalDepth, 1, totalDepth, ws.Dimension.End.Column].AutoFilter =
                        excelProtectionSetting.AutoFilterInCaseProtected;
                }

                if (excelProtectionSetting.AllowEditToAllData)
                {
                    // Unlock all cells in the worksheet
                    ws.Cells[totalDepth + 1, 1, ws.Dimension.End.Row, ws.Dimension.End.Column].Style.Locked = false;
                }
                else if (excelProtectionSetting.EditableColumnIndex != null &&
                         excelProtectionSetting.EditableColumnIndex.Count > 0)
                {
                    // Unlock only specified columns
                    foreach (var colIndex in excelProtectionSetting.EditableColumnIndex)
                    {
                        ws.Cells[totalDepth + 1, colIndex, ws.Dimension.End.Row, colIndex].Style.Locked = false;
                    }
                }

                // Now apply protection
                ws.Protection.IsProtected = true;
                ws.Protection.SetPassword(excelProtectionSetting.Password);

                // Disable insert/delete
                ws.Protection.AllowDeleteColumns = false;
                ws.Protection.AllowDeleteRows = false;
                ws.Protection.AllowInsertColumns = false;
                ws.Protection.AllowInsertRows = false;

                // Enable all other functionality
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

            // If header has subheaders, only span one row vertically; otherwise span full depth
            bool hasChildren = header.SubHeaders != null && header.SubHeaders.Any();
            int rowSpan = hasChildren ? 1 : totalDepth - usedDepth;
            int colSpan = leafCount;
            header.CollNo = col;
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

        private static void AddHiddenMetadataSheet(
            ExcelPackage package,
            string sheetName,
            IDictionary<string, string> metadata)
        {
            // Check if sheet already exists
            var existingSheet = package.Workbook.Worksheets.FirstOrDefault(ws => ws.Name == sheetName);
            if (existingSheet != null)
            {
                // Clear old content if needed
                //existingSheet.Cells.Clear();
                //existingSheet.Hidden = eWorkSheetHidden.Hidden;
                return;
            }

            // Create new hidden sheet
            var metadataSheet = package.Workbook.Worksheets.Add(sheetName);

            // metadata entries
            int row = 1;
            foreach (var kvp in metadata)
            {
                metadataSheet.Cells[row, 1].Value = kvp.Key;
                metadataSheet.Cells[row, 2].Value = kvp.Value;
                row++;
            }
            // Hide the sheet from user view
            metadataSheet.Hidden = eWorkSheetHidden.Hidden;
        }

        internal static void AddMetadataToSheet(
            ExcelPackage package,
            IDictionary<string, string> metadata,
            int dataStartRow,
            int dataEndRow,
            int dataStartCol,
            int dataEndCol)
        {
            //For metadata given by user, add it to a hidden sheet in the excel file
            if (metadata != null && metadata.Keys.Count > 0)
            {
                AddHiddenMetadataSheet(
                    package,
                    AppConstant.USER_DEFINED_METASHEET,
                    metadata);
            }

            //For metadata generated by the system, add it to a hidden sheet in the excel file
            AddHiddenMetadataSheet(
                package,
                AppConstant.EXCEL_HELPER_METASHEET,
                CreateExcelHelperMetaData(
                    dataStartRow,
                    dataEndRow,
                    dataStartCol,
                    dataEndCol));
        }

        private static IDictionary<string, string> CreateExcelHelperMetaData(
            int dataStartRow,
            int dataEndRow,
            int dataStartCol,
            int dataEndCol)
        {
            return new Dictionary<string, string>
            {
                { AppConstant.EXCEL_HELPER_METADATA_CREATED_ON_KEY, DateTime.UtcNow.ToString("u") },
                { AppConstant.EXCEL_HELPER_METADATA_AUTHOR_KEY, "ExcelService" },
                { AppConstant.EXCEL_HELPER_METADATA_DATA_ROW_STARTS_FROM, dataStartRow.ToString() },
                { AppConstant.EXCEL_HELPER_METADATA_DATA_ROW_END_TO, dataEndRow.ToString() },
                { AppConstant.EXCEL_HELPER_METADATA_DATA_COL_STARTS_FROM, dataStartCol.ToString() },
                { AppConstant.EXCEL_HELPER_METADATA_DATA_COL_END_TO, dataEndCol.ToString() },
            };
        }

        internal static IDictionary<string, string> ReadMetadataFromSheet(
            ExcelPackage package,
            string sheetName)
        {
            var result = new Dictionary<string, string>();
            // Find the sheet
            var metadataSheet = package.Workbook.Worksheets.FirstOrDefault(ws => ws.Name == sheetName);
            if (metadataSheet == null)
            {
                return result; // return empty if sheet not found
            }

            // Iterate rows until we hit an empty key cell
            int row = 1;
            while (true)
            {
                var key = metadataSheet.Cells[row, 1].Text;
                var value = metadataSheet.Cells[row, 2].Text;
                if (string.IsNullOrWhiteSpace(key))
                    break;
                result[key] = value;
                row++;
            }
            return result;
        }

        internal static T ConvertToType<T>(object value)
        {
            if (value == null || string.IsNullOrWhiteSpace(value.ToString()))
            {
                return GetDefaultValue<T>(typeof(T));
            }

            var targetType = typeof(T);

            // Handle Nullable<T>
            var underlyingType = Nullable.GetUnderlyingType(targetType);
            if (underlyingType != null)
            {
                return (T)ConvertToTypeDynamic(value, underlyingType);
            }

            return (T)ConvertToTypeDynamic(value, targetType);
        }

        private static object ConvertToTypeDynamic(object value, Type targetType)
        {
            // If already the right type, return directly
            if (value != null && targetType.IsAssignableFrom(value.GetType()))
            {
                return value;
            }

            string strValue = value?.ToString() ?? string.Empty;

            switch (Type.GetTypeCode(targetType))
            {
                case TypeCode.String:
                    return strValue;
                case TypeCode.Boolean:
                    return bool.TryParse(strValue, out var b) ? b : default(bool);
                case TypeCode.Byte:
                    return byte.TryParse(strValue, out var bt) ? bt : default(byte);
                case TypeCode.Char:
                    return char.TryParse(strValue, out var c) ? c : default(char);
                case TypeCode.Int16:
                    return short.TryParse(strValue, out var s) ? s : default(short);
                case TypeCode.UInt16:
                    return ushort.TryParse(strValue, out var us) ? us : default(ushort);
                case TypeCode.Int32:
                    return int.TryParse(strValue, out var i) ? i : default(int);
                case TypeCode.UInt32:
                    return uint.TryParse(strValue, out var ui) ? ui : default(uint);
                case TypeCode.Int64:
                    return long.TryParse(strValue, out var l) ? l : default(long);
                case TypeCode.UInt64:
                    return ulong.TryParse(strValue, out var ul) ? ul : default(ulong);
                case TypeCode.Single:
                    return float.TryParse(strValue, out var f) ? f : default(float);
                case TypeCode.Double:
                    return double.TryParse(strValue, out var d) ? d : default(double);
                case TypeCode.Decimal:
                    return decimal.TryParse(strValue, out var dec) ? dec : default(decimal);
                case TypeCode.DateTime:
                    return DateTime.TryParse(strValue, out var dt) ? dt : default(DateTime);
                case TypeCode.Object:
                    if (targetType == typeof(Guid))
                        return Guid.TryParse(strValue, out var g) ? g : Guid.Empty;

                    if (targetType == typeof(DateTimeOffset))
                        return DateTimeOffset.TryParse(strValue, out var dto) ? dto : default(DateTimeOffset);

                    if (targetType == typeof(TimeSpan))
                        return TimeSpan.TryParse(strValue, out var ts) ? ts : default(TimeSpan);

                    if (targetType.IsEnum)
                    {
                        try
                        {
                            return Enum.Parse(targetType, strValue, ignoreCase: true);
                        }
                        catch
                        {
                            return Activator.CreateInstance(targetType);
                        }
                    }
                    return Convert.ChangeType(value, targetType);
                default:
                    return Convert.ChangeType(value, targetType);
            }
        }

        private static T GetDefaultValue<T>(Type type)
        {
            if (!type.IsValueType)
            {
                return default(T);
            }
            return Activator.CreateInstance<T>();
        }

    }
}
