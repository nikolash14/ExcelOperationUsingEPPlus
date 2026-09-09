
using ExcelService.Helper;
using ExcelService.Model;
using ExcelService.Service;
using System;
using System.Collections.Generic;
using System.IO;

namespace TestApp
{
    internal class Program
    {
        static void Main(string[] args)
        {
            var excelBytes = ExcelServiceHelper.Export(HeaderLevel1(), DataLevel1());
            //var excelBytes = ExcelServiceHelper.Export(HeaderLevel3(), DataLevel3());
            // Target folder and file path
            string folderPath = @"D:\Temp\Reports";
            string filePath = Path.Combine(folderPath, "Orders.xlsx");

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
        private static List<ExcelHeader> NewSubHeader()
        {
            return new List<ExcelHeader>
                        {
                            new ExcelHeader { Name = "Qty", Style = ExcelServiceStyle.GoldAccent4L80 },
                            new ExcelHeader { Name = "Rate", Style = ExcelServiceStyle.GoldAccent4L80 }
                        };
        }
    }
}
