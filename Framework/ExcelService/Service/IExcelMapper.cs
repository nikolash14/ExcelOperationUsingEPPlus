using System.Collections.Generic;
using System.IO;

namespace ExcelService.Service
{
    public interface IExcelMapper
    {
        byte[] CreateExcel(
            string sheetName,
            List<Model.ExcelHeader> headers,
            Model.ExcelWorkbookStyle excelWorkbookStyle,
            IDictionary<string, string> metadata = null,
            Model.ExcelProtectionSetting excelProtectionSetting = null);

        Stream CreateExcelStream(
            string sheetName,
            List<Model.ExcelHeader> headers,
            Model.ExcelWorkbookStyle excelWorkbookStyle,
            IDictionary<string, string> metadata = null,
            Model.ExcelProtectionSetting excelProtectionSetting = null);

        IReadOnlyList<IReadOnlyList<T>> ReadExcelData<T>(
            byte[] excelData,
            string sheetName,
            IDictionary<string, string> metadata = null);
    }
}
