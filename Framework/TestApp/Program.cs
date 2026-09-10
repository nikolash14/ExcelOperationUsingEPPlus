
using ExcelService.Helper;
using ExcelService.Model;
using ExcelService.Service;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;

namespace TestApp
{
    internal class Program
    {
        static void Main(string[] args)
        {
            CreateExcel();
            ReadExcel();
        }
        private static void ReadExcel()
        {
            string filePath = @"D:\Temp\Reports\Orders3.json";
            var excelBytes = File.ReadAllBytes(@"D:\Temp\Reports\Orders3.xlsx");
            var excelData = ExcelServiceHelper.ReadExcelData<ExcelHeader>(
                excelBytes, "excelSheet");
            if (excelData == null)
                throw new ArgumentNullException(nameof(excelData));

            if (string.IsNullOrWhiteSpace(filePath))
                throw new ArgumentException("File path cannot be null or empty.", nameof(filePath));

            // Serialize with indentation for readability
            var options = new JsonSerializerOptions
            {
                WriteIndented = true
            };

            string json = JsonSerializer.Serialize(excelData, options);

            File.WriteAllText(filePath, json);
        }

        private static void CreateExcel()
        {
            var workbookStyle = new ExcelWorkbookStyle
            {
                ApplyBorders = true,
                AutoFitColumns = true,
                FreezeHeader = true,
                FreezeColumnCount = 3
            };
            var excelProtectionSetting = new ExcelProtectionSetting
            {
                IsProtected = true,
                Password = "excelProtectionSetting1234",
                AutoFilterInCaseProtected = false,
                AllowEditToAllData = false,
                EditableColumnIndex = new List<int> { 1, 2, 4 }
            };
            WriteExcelToFile(
                ExcelServiceHelper.CreateExcel(
                        "excelSheet",
                        HeaderLevel1(),
                        DataLevel1(),
                        workbookStyle,
                        null,
                        excelProtectionSetting
                        ),
                "Orders1.xlsx");
            WriteExcelToFile(
            ExcelServiceHelper.CreateExcel(
                    "excelSheet",
                    HeaderLevel2(),
                    DataLevel2(),
                    workbookStyle,
                    null,
                    excelProtectionSetting),
            "Orders2.xlsx");
            WriteExcelToFile(
                ExcelServiceHelper.CreateExcel(
                        "excelSheet",
                        HeaderLevel3(),
                        DataLevel3(),
                        workbookStyle,
                        null,
                        excelProtectionSetting),
                "Orders3.xlsx");
        }

        private static void WriteExcelToFile(byte[] excelBytes, string fileName)
        {
            string folderPath = @"D:\Temp\Reports";
            string filePath = Path.Combine(folderPath, fileName);
            // Ensure folder exists
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }
            // Write byte array to file (creates new file if not exists, overwrites if exists)
            File.WriteAllBytes(filePath, excelBytes);
            Console.WriteLine("Excel file created successfully at: " + filePath);
        }

        private static List<List<object>> DataLevel3()
        {
            return new List<List<object>>
            {
                new List<object> { "A1", "B1", "C1", 3, 8.42, 0, 8.57, 1, 6.31, false },
                new List<object> { "A2", "B2", "C2", 1, 8.77, 3, 7.96, 2, 6.30, false },
                new List<object> { "A3", "B3", "C3", 1, 8.46, 0, 9.12, 4, 7.70, true }
            };
        }
        private static List<List<object>> DataLevel1()
        {
            return new List<List<object>>
            {
                new List<object> { "A1", "B1", "C1", 3, 8.42, false },
                new List<object> { "A2", "B2", "C2", 1, 8.77, false },
                new List<object> { "A3", "B3", "C3", 1, 8.46, true }
            };
        }
        private static List<ExcelHeader> HeaderLevel1()
        {
            return new List<ExcelHeader>
                {
                    new ExcelHeader { Name = "Order1", Style = ExcelServiceStyle.BlueAccent5L80 },
                    new ExcelHeader { Name = "Order2", Style = ExcelServiceStyle.BlueAccent5L80 },
                    new ExcelHeader { Name = "Order3", Style = ExcelServiceStyle.BlueAccent5L80 },
                    new ExcelHeader { Name = "Qty" , Style = ExcelServiceStyle.GreenAccent6L80},
                    new ExcelHeader { Name = "Rate", Style = ExcelServiceStyle.GreenAccent6L80 },
                    new ExcelHeader
                    {
                        Name = "Term & Condition",
                        Style = ExcelServiceStyle.GoldAccent4L80
                    }
                };
        }
        private static List<ExcelHeader> HeaderLevel3()
        {
            return new List<ExcelHeader>
        {
            new ExcelHeader
            {
                Name = "Order",
                Style = ExcelServiceStyle.BlueAccent5L60,
                SubHeaders = new List<ExcelHeader>
                {
                    new ExcelHeader { Name = "Order1", Style = ExcelServiceStyle.BlueAccent5L80 },
                    new ExcelHeader { Name = "Order2", Style = ExcelServiceStyle.BlueAccent5L80 },
                    new ExcelHeader { Name = "Order3", Style = ExcelServiceStyle.BlueAccent5L80 }
                }
            },
            new ExcelHeader
            {
                Name = "Charges",
                Style = ExcelServiceStyle.GoldAccent4L40,
                SubHeaders = new List<ExcelHeader>
                {
                    new ExcelHeader
                    {
                        Name = "V1",
                        Style= ExcelServiceStyle.GoldAccent4L60,
                        SubHeaders = NewSubHeader()
                    },
                    new ExcelHeader
                    {
                        Name = "V2",
                        Style= ExcelServiceStyle.GoldAccent4L60,
                        SubHeaders = NewSubHeader()
                    },
                    new ExcelHeader
                    {
                        Name = "V3",
                        Style= ExcelServiceStyle.GoldAccent4L60,
                        SubHeaders = NewSubHeader()
                    }
                }
            },
            new ExcelHeader
            {
                Name = "Term & Condition",
                Style = ExcelServiceStyle.GreenAccent6L80
            }
        };
        }

        private static List<List<object>> DataLevel2()
        {
            return new List<List<object>>
            {
                new List<object> { "A1", "B1", "C1", 3, 8.42, 0, 8.57, 1, 6.31},
                new List<object> { "A2", "B2", "C2", 1, 8.77, 3, 7.96, 2, 6.30},
                new List<object> { "A3", "B3", "C3", 1, 8.46, 0, 9.12, 4, 7.70 }
            };
        }
        private static List<ExcelHeader> HeaderLevel2()
        {
            return new List<ExcelHeader>
        {
            new ExcelHeader
            {
                Name = "Order",
                Style = ExcelServiceStyle.BlueAccent5L60,
                SubHeaders = new List<ExcelHeader>
                {
                    new ExcelHeader { Name = "Order1", Style = ExcelServiceStyle.BlueAccent5L80 },
                    new ExcelHeader { Name = "Order2", Style = ExcelServiceStyle.BlueAccent5L80 },
                    new ExcelHeader { Name = "Order3", Style = ExcelServiceStyle.BlueAccent5L80 }
                }
            },
            new ExcelHeader
            {
                Name = "V1",
                Style= ExcelServiceStyle.GoldAccent4L60,
                SubHeaders = NewSubHeader()
            },
            new ExcelHeader
            {
                Name = "V2",
                Style= ExcelServiceStyle.GoldAccent4L60,
                SubHeaders = NewSubHeader()
            },
            new ExcelHeader
            {
                Name = "V3",
                Style= ExcelServiceStyle.GoldAccent4L60,
                SubHeaders = NewSubHeader()
            }
        };
        }


        private static List<ExcelHeader> NewSubHeader()
        {
            return new List<ExcelHeader>
                        {
                            new ExcelHeader { Name = "Qty", Style = ExcelServiceStyle.GoldAccent4L80 },
                            new ExcelHeader { Name = "Rate", Style = ExcelServiceStyle.GoldAccent4L80 }
                        };
        }
    }

    public class ExcelRowData
{
    public string ColA { get; set; }   // "A1"
    public string ColB { get; set; }   // "B1"
    public string ColC { get; set; }   // "C1"
    public int ColD { get; set; }      // 3
    public double ColE { get; set; }   // 8.42
    public int ColF { get; set; }      // 0
    public double ColG { get; set; }   // 8.57
    public int ColH { get; set; }      // 1
    public double ColI { get; set; }   // 6.31
    public bool ColJ { get; set; }     // false
}

}
