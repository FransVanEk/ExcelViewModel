using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WannaApp.Excel.Attributes;
using WannaApp.Excel.Helpers;

namespace WannaApp.ExcelViewModelTests.TestObjecten
{
    public class NewTestObject
    {
        [ExcelMappingKey]
        [ExcelMappingName("Eerste Sleutel" , OrderIndex=1) ]
        public int FirstKey { get; set; }

        [ExcelMappingKey]
        [ExcelMappingName("Tweede sleutel" , OrderIndex  =2)]
        public string SecondKey { get; set; }
        
        [ExcelMappingName("Getal", OrderIndex = 4)]
        public int Int { get; set; }
       
        [ExcelMappingIgnore]
        public bool Ignore { get; set; }

        [ExcelMappingName("Tekst", OrderIndex=3)]
        public string String { get; set; }

        [ExcelMappingName("Datum")]
        public DateTime DateTime { get; set; }

        [ExcelMappingName("Decimaal")]
        public double Double { get; set; }

        public List<string> DynamicStrings { get; set; }
       
        public List<int> DynamicInts { get; set; }

        public Guid Guid { get; set;}

    }
}
