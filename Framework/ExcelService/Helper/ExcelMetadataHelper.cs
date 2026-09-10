using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;

namespace ExcelService.Helper
{
    internal static class ExcelMetadataHelper
    {
        private static void AddHiddenMetadataSheet(
            ExcelPackage package,
            string sheetName,
            IDictionary<string, string> metadata)
        {
            var existingSheet = package.Workbook.Worksheets.FirstOrDefault(ws => ws.Name == sheetName);
            if (existingSheet != null)
            {
                return;
            }

            var metadataSheet = package.Workbook.Worksheets.Add(sheetName);
            int row = 1;
            foreach (var kvp in metadata)
            {
                metadataSheet.Cells[row, 1].Value = kvp.Key;
                metadataSheet.Cells[row, 2].Value = kvp.Value;
                row++;
            }
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
            if (metadata != null && metadata.Keys.Count > 0)
            {
                AddHiddenMetadataSheet(package, AppConstant.USER_DEFINED_METASHEET, metadata);
            }

            AddHiddenMetadataSheet(package, AppConstant.EXCEL_HELPER_METASHEET, CreateExcelHelperMetaData(dataStartRow, dataEndRow, dataStartCol, dataEndCol));
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
            var metadataSheet = package.Workbook.Worksheets.FirstOrDefault(ws => ws.Name == sheetName);
            if (metadataSheet == null)
            {
                return result;
            }

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
    }
}
