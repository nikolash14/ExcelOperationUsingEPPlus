using System.Collections.Generic;
using System.Linq;
namespace ExcelService.Model
{
    public class ExcelHeader
    {
        public string Name { get; set; }
        public ExcelStyle Style { get; set; }
        public List<ExcelHeader> SubHeaders { get; set; }

        public ExcelHeader()
        {
            Name = string.Empty;
            Style = new ExcelStyle();
            SubHeaders = new List<ExcelHeader>();
        }

        public int GetDepth()
        {
            if (SubHeaders == null || SubHeaders.Count == 0)
                return 1;

            return 1 + SubHeaders.Max(h => h.GetDepth());
        }

        public int GetLeafCount()
        {
            if (SubHeaders == null || SubHeaders.Count == 0)
                return 1;
            return SubHeaders.Sum(h => h.GetLeafCount());
        }
    }

}
