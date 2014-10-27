using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WannaApp.Excel.Attributes
{
    [AttributeUsage(AttributeTargets.Property | AttributeTargets.Method, AllowMultiple = false)]
    public class ExcelMappingKeyAttribute : Attribute
    {
    }
}
