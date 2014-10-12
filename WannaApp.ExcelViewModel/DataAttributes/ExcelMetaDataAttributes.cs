using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WannaApp.ExcelViewModel.DataAttributes
{

    [System.AttributeUsage(AttributeTargets.Property)]
    public class Column : Attribute
    {

        public Column(string displayColumnName)
        {
            this.DisplayColumnName = displayColumnName;
        }
        
        public bool Hidden { get; set; }
        public string DisplayColumnName { get; set; }

        public int OrderIndex { get; set; }
        public bool Autowidth { get; set; }
        public int ManualWidth { get; set; }
    }
}
