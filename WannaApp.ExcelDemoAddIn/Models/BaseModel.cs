using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WannaApp.Excel.Helpers;

namespace WannaApp.ExcelDemoAddIn.Models
{
    public class BaseModel
    {
        [ExcelMappingKey]
        [ExcelMappingName("Eerste Sleutel", OrderIndex = 10)]
        public int FirstKey { get; set; }

        [ExcelMappingKey]
        [ExcelMappingName("Tweede sleutel", OrderIndex = 20)]
        public string SecondKey { get; set; }

        [ExcelMappingName("Getal", OrderIndex = 40)]
        public int Int { get; set; }

        [ExcelMappingIgnore]
        public bool Ignore { get; set; }

        [ExcelMappingName("Tekst", OrderIndex = 30)]
        public string String { get; set; }

        [ExcelMappingName("Datum")]
        public DateTime DateTime { get; set; }

        [ExcelMappingName("Decimaal")]
        public double Double { get; set; }

        [ExcelMappingName("dede", OrderIndex =25)]
        public List<string> DynamicStrings { get; set; }

        public List<int> DynamicInts { get; set; }

    }
}
