using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WannaApp.Excel.Helpers
{
    public class HeadersMappingInfo
    {
        public string PropertyName { get; set; }
        public string HeaderText { get; set; }
        public bool IsDynamicRangeProperty { get; set; }
        public int ColumnIndex { get; set; }

    }
}
