using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WannaApp.Excel.Helpers
{
    [AttributeUsage(AttributeTargets.Property | AttributeTargets.Method, AllowMultiple = false)]
    public class ExcelMappingIgnoreAttribute : Attribute
    {
        // Empty
    }
}
