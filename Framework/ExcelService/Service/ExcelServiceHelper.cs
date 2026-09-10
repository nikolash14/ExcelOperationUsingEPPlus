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
    public class ExcelServiceHelper
    {
        // Returns the maximum depth (number of header rows) required for the provided headers
        public static int GetTotalDepth(List<ExcelHeader> headers)
        {
            if (headers == null || headers.Count == 0)
                return 0;
            return headers.Max(h => h.GetDepth());
        }

        public static byte[] CreateExcel<T>(
            string sheetName,
            List<ExcelHeader> headers,
            List<List<T>> data,
            ExcelWorkbookStyle excelWorkbookStyle,
            IDictionary<string, string> metadata = null,
            ExcelProtectionSetting excelProtectionSetting = null)
        {
            int totalDepth = GetTotalDepth(headers);
            ExcelPackage.License.SetNonCommercialOrganization(AppConstant.EXCEL_HELPER_LICENSING_ORGANIZATION);
            using (var package = new ExcelPackage())
            {
                var ws = package.Workbook.Worksheets.Add(sheetName);
                // Render headers
                int col = 1;
                foreach (var header in headers)
                {
                    ExcelHelper.RenderHeaders(ws, header, 1, col, totalDepth);
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

                //Auto-fit columns if requested
                ExcelHelper.AutoFitColumns(
                    excelWorkbookStyle,
                    ws);

                // ... inside ExcelServiceHelper.Export, after the data rows loop
                //int dataStartRow = headers.Max(h => h.GetDepth()) + 1;
                // Apply thin borders to every cell in the data area
                int lastRow = row - 1; // row was incremented after last record
                int lastCol = data.Count > 0 ? data[0].Count : headers.Sum(h => h.GetLeafCount());
                ExcelHelper.ApplyBorders(
                    excelWorkbookStyle,
                    ws,
                    lastRow,
                    lastCol);

                // Freeze panes if requested
                ExcelHelper.FreezePanes(
                    excelWorkbookStyle,
                    totalDepth,
                    ws);

                // Add metadata to a hidden sheet if provided
                ExcelHelper.AddMetadataToSheet(
                    package,
                    metadata,
                    totalDepth + 1,
                    lastRow,
                    1,
                    lastCol);

                //Excel Protection Settings
                ExcelHelper.ApplyExcelProtection(
                    excelProtectionSetting,
                    package,
                    ws,
                    totalDepth);

                //Return the Excel file as a byte array
                return package.GetAsByteArray();
            }
        }

        public static IReadOnlyList<IReadOnlyList<T>> ReadExcelData<T>(
            byte[] excelData,
            string sheetName,
            IDictionary<string, string> metadata = null)
        {
            var result = new List<IReadOnlyList<T>>();
            using (var stream = new MemoryStream(excelData))
            {
                ExcelPackage.License.SetNonCommercialOrganization(AppConstant.EXCEL_HELPER_LICENSING_ORGANIZATION);
                using (var package = new ExcelPackage(stream))
                {
                    var worksheet = package.Workbook.Worksheets[sheetName];
                    if (worksheet == null)
                        throw new ArgumentException($"Worksheet '{sheetName}' not found.");
                    // Validate the metadata if provided with the metadata in the excel file
                    // this meta data is provided by the user when the excel file was created and saved in a hidden sheet in the excel file
                    // If the provided metadata does not match the metadata in the excel file, throw an exception
                    if (metadata != null && metadata.Any())
                    {
                        var excelMetadata = ExcelHelper.ReadMetadataFromSheet(package, AppConstant.USER_DEFINED_METASHEET);
                        foreach (var kvp in metadata)
                        {
                            if (!excelMetadata.ContainsKey(kvp.Key) || excelMetadata[kvp.Key] != kvp.Value)
                            {
                                throw new ArgumentException($"Metadata mismatch for key '{kvp.Key}'.");
                            }
                        }
                    }

                    // Get the how much to be read from the worksheet
                    // based on the dimension of the worksheet that was saved in meta data
                    var dataReaderMetaData = ExcelHelper.ReadMetadataFromSheet(package, AppConstant.EXCEL_HELPER_METASHEET);
                    var startRow = int.Parse(dataReaderMetaData[AppConstant.EXCEL_HELPER_METADATA_DATA_ROW_STARTS_FROM]);
                    var endRow = int.Parse(dataReaderMetaData[AppConstant.EXCEL_HELPER_METADATA_DATA_ROW_END_TO]);
                    var startCol = int.Parse(dataReaderMetaData[AppConstant.EXCEL_HELPER_METADATA_DATA_COL_STARTS_FROM]);
                    var endCol = int.Parse(dataReaderMetaData[AppConstant.EXCEL_HELPER_METADATA_DATA_COL_END_TO]);

                    // add: using System.Reflection;
                    for (int row = startRow; row <= endRow; row++)
                    {
                        // If T is a complex class (not string), construct one T per row by mapping columns to writable properties
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
                            // add the constructed object as the single-item row
                            result.Add(new List<T> { instance });
                        }
                        else
                        {
                            // primitive or string: keep existing per-cell behavior
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
