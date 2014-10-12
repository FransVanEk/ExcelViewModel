using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WannaApp.Excel.Helpers;


namespace WannaApp.ExcelDemoAddIn.Models
{
    internal class Demo
    {
        [ExcelMappingName("Name")]
        public string Naam { get; set; }

        [ExcelMappingName("key")]
        public string sleutel { get; set; }

    }
}
