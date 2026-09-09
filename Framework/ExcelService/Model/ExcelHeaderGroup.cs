namespace ExcelService.Model
{
    public class ExcelHeaderGroup
    {
        public string GroupName { get;}
        public ExcelStyle ExcelStyle { get; }
        public ExcelHeader ExcelHeader { get; }

        public ExcelHeaderGroup(string groupName, ExcelHeader excelHeader, ExcelStyle excelStyle = null)
        {
            GroupName = groupName;
            ExcelHeader = excelHeader;
            ExcelStyle = excelStyle ?? new ExcelStyle();
        }
    }
}
