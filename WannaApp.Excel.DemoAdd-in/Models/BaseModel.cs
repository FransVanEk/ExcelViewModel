using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WannaApp.Excel.Attributes;

namespace WannaApp.Excel.DemoAdd_in.Models
{
  
        public class BaseModel
        {
            [ExcelMappingKey]
            [ExcelMappingName("First Key", OrderIndex = 10)]
            public int FirstKey { get; set; }

            [ExcelMappingKey]
            [ExcelMappingName("Second Key", OrderIndex = 20)]
            public string SecondKey { get; set; }

            [ExcelMappingName("number", OrderIndex = 40)]
            public int Int { get; set; }

            [ExcelMappingIgnore]
            public bool Ignore { get; set; }

            [ExcelMappingName("Text", OrderIndex = 30)]
            public string String { get; set; }

            [ExcelMappingName("Date")]
            public DateTime DateTime { get; set; }

            [ExcelMappingName("Decimal")]
            public double Double { get; set; }

            public List<string> DynamicStrings { get; set; }

            public List<int> DynamicInts { get; set; }

        }
    }

