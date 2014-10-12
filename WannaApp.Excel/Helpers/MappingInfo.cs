using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace WannaApp.Excel.Helpers
{
    public class MappingInfo
    {
        public int ColumnIndex { get; set; }
        public string ColumnName { get; set; }
        public string PropertyName { get; set; }
        public bool IsDynamicRange { get { return (DynamicColumnNames != null && DynamicColumnNames.Count > 0); } }
        public bool IsKey { get; set; }
        public Type Type { get; set; }
        public List<string> DynamicColumnNames { get; set; }

    }
}
