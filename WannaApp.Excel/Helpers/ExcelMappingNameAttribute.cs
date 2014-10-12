using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WannaApp.Excel.Helpers
{
    [AttributeUsage(AttributeTargets.Property | AttributeTargets.Method, AllowMultiple = false)]
    public class ExcelMappingNameAttribute : Attribute
    {
        public string Name { get; set; }
        public int OrderIndex { get; set; }

      

        public ExcelMappingNameAttribute()
        {
            //Empty
            SetDefaults();
        }

        public ExcelMappingNameAttribute(string name)
        {
            Name = name;
            SetDefaults();
        }

        public ExcelMappingNameAttribute(Type resourceType, string resourceName)
        {
            Name = ResourceHelper.GetResourceLookup<string>(resourceType, resourceName);

            SetDefaults();
        }

        private void SetDefaults()
        {
            this.OrderIndex = int.MaxValue;
        }
    }
}
