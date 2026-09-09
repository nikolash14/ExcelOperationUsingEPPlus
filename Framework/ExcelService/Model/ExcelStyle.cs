using OfficeOpenXml.Style;

namespace ExcelService.Model
{
    public class ExcelStyle
    {
        public bool Bold { get; set; } = true;
        public int FontSize { get; set; } = 11;
        public ExcelHorizontalAlignment HorizontalAlignment { get; set; } = ExcelHorizontalAlignment.Center;
        public ExcelVerticalAlignment VerticalAlignment { get; set; } = ExcelVerticalAlignment.Center;
        public System.Drawing.Color FontColor { get; set; } = System.Drawing.Color.Black;
        public System.Drawing.Color BackgroundColor { get; set; } = System.Drawing.Color.White;
    }
}
