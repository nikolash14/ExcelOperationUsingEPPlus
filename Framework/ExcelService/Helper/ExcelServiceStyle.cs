using ExcelService.Model;

namespace ExcelService.Helper
{
    public class ExcelServiceStyle
    {
        public static ExcelStyle DefaultHeaderStyle => new ExcelStyle
        {
            Bold = true,
            FontSize = 11,
            HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center,
            VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center,
            FontColor = System.Drawing.Color.Black,
            BackgroundColor = System.Drawing.Color.FromArgb(180, 198, 231)
        };

        #region GreenAccent6
        public static ExcelStyle GreenAccent6L40 => new ExcelStyle
        {
            Bold = true,
            FontSize = 11,
            HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center,
            VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center,
            FontColor = System.Drawing.Color.Black,
            BackgroundColor = System.Drawing.Color.FromArgb(169, 208, 142)
        };
        public static ExcelStyle GreenAccent6L60 => new ExcelStyle
        {
            Bold = true,
            FontSize = 11,
            HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center,
            VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center,
            FontColor = System.Drawing.Color.Black,
            BackgroundColor = System.Drawing.Color.FromArgb(198, 224, 180)
        };
        public static ExcelStyle GreenAccent6L80 => new ExcelStyle
        {
            Bold = true,
            FontSize = 11,
            HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center,
            VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center,
            FontColor = System.Drawing.Color.Black,
            BackgroundColor = System.Drawing.Color.FromArgb(226, 239, 218)
        };
        #endregion

        #region GoldAccent4
        public static ExcelStyle GoldAccent4L40 => new ExcelStyle
        {
            Bold = true,
            FontSize = 11,
            HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center,
            VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center,
            FontColor = System.Drawing.Color.Black,
            BackgroundColor = System.Drawing.Color.FromArgb(234, 200, 82)
        };

        public static ExcelStyle GoldAccent4L60 => new ExcelStyle
        {
            Bold = true,
            FontSize = 11,
            HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center,
            VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center,
            FontColor = System.Drawing.Color.Black,
            BackgroundColor = System.Drawing.Color.FromArgb(238, 213, 115)
        };
        public static ExcelStyle GoldAccent4L80 => new ExcelStyle
        {
            Bold = true,
            FontSize = 11,
            HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center,
            VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center,
            FontColor = System.Drawing.Color.Black,
            BackgroundColor = System.Drawing.Color.FromArgb(242, 226, 149)
        };
        #endregion

        #region BlueAccent5
        public static ExcelStyle BlueAccent5L40 => new ExcelStyle
        {
            Bold = true,
            FontSize = 11,
            HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center,
            VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center,
            FontColor = System.Drawing.Color.Black,
            BackgroundColor = System.Drawing.Color.FromArgb(180, 198, 231)
        };
        public static ExcelStyle BlueAccent5L60 => new ExcelStyle
        {
            Bold = true,
            FontSize = 11,
            HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center,
            VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center,
            FontColor = System.Drawing.Color.Black,
            BackgroundColor = System.Drawing.Color.FromArgb(142, 169, 219)
        };
        public static ExcelStyle BlueAccent5L80 => new ExcelStyle
        {
            Bold = true,
            FontSize = 11,
            HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center,
            VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center,
            FontColor = System.Drawing.Color.Black,
            BackgroundColor = System.Drawing.Color.FromArgb(189, 215, 238)
        }; 
        #endregion

        public static ExcelStyle DefaultDataStyle => new ExcelStyle
        {
            Bold = false,
            FontSize = 11,
            HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left,
            VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center,
            FontColor = System.Drawing.Color.Black,
            BackgroundColor = System.Drawing.Color.White
        };

    }
}
