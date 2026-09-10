using System;
using System.Collections.Generic;

namespace ExcelService.Model
{
    public class ExcelProtectionSetting
    {
        public Boolean IsProtected { get; set; } = false;
        public string Password { get; set; } = "ExcelProtectionSetting@123";
        public Boolean AutoFilterInCaseProtected { get; set; } = true;
        public Boolean AllowEditToAllData { get; set; } = true;
        public List<int> EditableColumnIndex { get; set; }
    }
}
