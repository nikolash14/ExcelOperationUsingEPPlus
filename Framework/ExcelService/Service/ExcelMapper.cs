using ExcelService.Helper;
using ExcelService.Model;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;

namespace ExcelService.Service
{
    public class ExcelMapper : IExcelMapper
    {
        // Instance implementation to improve testability and DI
        public ExcelMapper()
        {
        }

        public byte[] CreateExcel(
            string sheetName,
            List<ExcelHeader> headers,
            ExcelWorkbookStyle excelWorkbookStyle,
            IDictionary<string, string> metadata = null,
            ExcelProtectionSetting excelProtectionSetting = null)
        {
            using (var ms = CreateExcelStream(sheetName, headers, excelWorkbookStyle, metadata, excelProtectionSetting) as MemoryStream)
            {
                return ms?.ToArray() ?? Array.Empty<byte>();
            }
        }

        public Stream CreateExcelStream(
            string sheetName,
            List<ExcelHeader> headers,
            ExcelWorkbookStyle excelWorkbookStyle,
            IDictionary<string, string> metadata = null,
            ExcelProtectionSetting excelProtectionSetting = null)
        {
            if (string.IsNullOrWhiteSpace(sheetName))
                throw new ArgumentException("sheetName must be provided", nameof(sheetName));

            if (headers == null)
                throw new ArgumentNullException(nameof(headers));

            int totalDepth = headers.Any() ? headers.Max(h => h.GetDepth()) : 0;

            // EPPlus license must be initialized externally via ExcelLicenseInitializer.Initialize()
            ExcelLicenseInitializer.Initialize();
            var package = new ExcelPackage();
            var ws = package.Workbook.Worksheets.Add(sheetName);

            // Render headers
            int col = 1;
            foreach (var header in headers)
            {
                ExcelRenderer.RenderHeaders(ws, header, 1, col, totalDepth);
                col += header.GetLeafCount();
            }

            // Render data rows
            int row = headers.Any() ? headers.Max(h => h.GetDepth()) + 1 : 1;
            var rowData = row;
            foreach (var record in headers)
            {
                rowData = row;
                foreach (var value in record.Data)
                {
                    ws.Cells[rowData, record.CollNo, rowData, record.CollNo].Value = value;
                    rowData++;
                }
            }
            row = rowData;

            //Auto-fit columns if requested
            ExcelRenderer.AutoFitColumns(excelWorkbookStyle, ws);

            // Apply thin borders to every cell in the data area
            int lastRow = row - 1; // row was incremented after last record
            int lastCol = headers.Sum(h => h.GetLeafCount());
            ExcelRenderer.ApplyBorders(excelWorkbookStyle, ws, lastRow, lastCol);

            // Freeze panes if requested
            ExcelRenderer.FreezePanes(excelWorkbookStyle, totalDepth, ws);

            // Add metadata to a hidden sheet if provided
            ExcelMetadataHelper.AddMetadataToSheet(package, metadata, totalDepth + 1, lastRow, 1, lastCol);

            //Excel Protection Settings
            ExcelRenderer.ApplyExcelProtection(excelProtectionSetting, package, ws, totalDepth);

            var ms = new MemoryStream();
            package.SaveAs(ms);
            ms.Position = 0;
            return ms;
        }

        public IReadOnlyList<IReadOnlyList<T>> ReadExcelData<T>(
            byte[] excelData,
            string sheetName,
            IDictionary<string, string> metadata = null)
        {
            var result = new List<IReadOnlyList<T>>();
            using (var stream = new MemoryStream(excelData))
            {
                using (var package = new ExcelPackage(stream))
                {
                    var worksheet = package.Workbook.Worksheets[sheetName];
                    if (worksheet == null)
                        throw new ArgumentException($"Worksheet '{sheetName}' not found.");

                    // Validate the metadata if provided
                    if (metadata != null && metadata.Any())
                    {
                        var excelMetadata = ExcelMetadataHelper.ReadMetadataFromSheet(package, AppConstant.USER_DEFINED_METASHEET);
                        foreach (var kvp in metadata)
                        {
                            if (!excelMetadata.ContainsKey(kvp.Key) || excelMetadata[kvp.Key] != kvp.Value)
                            {
                                throw new ArgumentException($"Metadata mismatch for key '{kvp.Key}'.");
                            }
                        }
                    }

                    var dataReaderMetaData = ExcelMetadataHelper.ReadMetadataFromSheet(package, AppConstant.EXCEL_HELPER_METASHEET);
                    var startRow = int.Parse(dataReaderMetaData[AppConstant.EXCEL_HELPER_METADATA_DATA_ROW_STARTS_FROM]);
                    var endRow = int.Parse(dataReaderMetaData[AppConstant.EXCEL_HELPER_METADATA_DATA_ROW_END_TO]);
                    var startCol = int.Parse(dataReaderMetaData[AppConstant.EXCEL_HELPER_METADATA_DATA_COL_STARTS_FROM]);
                    var endCol = int.Parse(dataReaderMetaData[AppConstant.EXCEL_HELPER_METADATA_DATA_COL_END_TO]);

                    for (int row = startRow; row <= endRow; row++)
                    {
                        if (typeof(T).IsClass && typeof(T) != typeof(string))
                        {
                            var props = typeof(T).GetProperties().Where(p => p.CanWrite).ToArray();
                            var instance = Activator.CreateInstance<T>();
                            var convertMethod = typeof(ExcelHelper).GetMethod("ConvertToTypeDynamic", BindingFlags.NonPublic | BindingFlags.Static);

                            for (int col = startCol; col <= endCol && (col - startCol) < props.Length; col++)
                            {
                                var cell = worksheet.Cells[row, col];
                                var rawValue = cell.Value;
                                if (rawValue == null || string.IsNullOrWhiteSpace(rawValue.ToString()))
                                    continue;

                                var prop = props[col - startCol];
                                var converted = convertMethod.Invoke(null, new object[] { rawValue, prop.PropertyType });
                                prop.SetValue(instance, converted);
                            }
                            result.Add(new List<T> { instance });
                        }
                        else
                        {
                            var rowData = new List<T>();
                            for (int col = startCol; col <= endCol; col++)
                            {
                                var cell = worksheet.Cells[row, col];
                                var rawValue = cell.Value;

                                if (rawValue == null || string.IsNullOrWhiteSpace(rawValue?.ToString()))
                                {
                                    rowData.Add(default);
                                }
                                else
                                {
                                    rowData.Add(ExcelHelper.ConvertToType<T>(rawValue));
                                }
                            }
                            result.Add(rowData);
                        }
                    }
                }
            }
            return result;
        }
    }
}
