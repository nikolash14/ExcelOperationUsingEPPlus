using System;

namespace ExcelService.Model
{
    public class ExcelWorkbookStyle
    {
        public Boolean ApplyBorders { get; set; } = true;
        public Boolean AutoFitColumns { get; set; } = true;
        public Boolean FreezeHeader { get; set; } = true;
        public int FreezeColumnCount { get; set; } = 0;

    }
}
