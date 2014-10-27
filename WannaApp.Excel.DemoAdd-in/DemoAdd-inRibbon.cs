using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using WannaApp.Excel.DataObjects;
using WannaApp.Excel.ExcelObjects;
using WannaApp.Excel.Extensions;
using WannaApp.Excel.Helpers.MappingHelpers;
using System.Drawing; 

namespace WannaApp.Excel.DemoAdd_in
{
    public partial class DemoAdd_inRibbon
    {
        private void DemoAdd_inRibbon_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void btn_BasicUsage_LoadIntoExcel_Click(object sender, RibbonControlEventArgs e)
        {
            var workbook = new ExcelWorkbook(Globals.ThisAddIn.Application.ActiveWorkbook);

            var list = workbook.FindOrCreateWorksheet("Basic_Usage").CreateListObject("A1", GetData(), "First_ListObject");
            var list2 = workbook.FindOrCreateWorksheet("Basic_Usage").CreateListObject("G1", GetData(), "Second_listObject");
            var list3 = workbook.FindOrCreateWorksheet("Basic_Usage_2").CreateListObject("A1", GetData(), "Third__ListObject");

            list.TableStyle("TableStyleLight12");
        }

        private ListObjectDataObject GetData()
        {
            var result = new ListObjectDataObject();
            result.HeaderValues = new string[] { "a", "b", "c", "d" };
            result.DataValues = new object[100,4];

            for (int i = 0; i < 100; i++)
            {
                var currentItem = new List<string>();
                for (int y = 0; y < 4; y++)
                {
                    result.DataValues[i,y] = ((int)y+i).ToString();
                }
            }
            return result;
        }

        private void btn_LoadObjectDataIntoExcel_Click(object sender, RibbonControlEventArgs e)
        {

            var workbook = new ExcelWorkbook(Globals.ThisAddIn.Application.ActiveWorkbook);
            var helper = GetHelperForBaseModel()
                .TransferToExcelFormat(GetTestDataNewTestObjects());

            var listobject = workbook.FindOrCreateWorksheet("LoadObjectData").CreateListObject("A1", helper, "FirstObjectList");
            listobject.GetHeaderRange().Orientation(45).Font(13);
            var test = listobject.GetHeaderRange().GetExcelRangeFor("First Dynamic", "Second Dynamic").Group().BackgroundColor(Color.LightBlue);

        }

         private TransferHelper<Models.BaseModel> GetHelperForBaseModel()
        {
            var helper =  new Helpers.MappingHelpers.TransferHelper<Models.BaseModel>()
               .SetDynamicColumnsFor("DynamicStrings", new List<string> { "First Dynamic", "Second Dynamic" })
               .SetDynamicColumnsFor("DynamicInts", new List<string> { "First", "second" });
            return helper;
        }


        private IEnumerable<Models.BaseModel> GetTestDataNewTestObjects()
        {
            var result = new List<Models.BaseModel>();
            result.Add(new Models.BaseModel { DateTime = DateTime.Now, Double = 1.5, FirstKey = 1, Ignore = false, Int = 1, SecondKey = "second", String = "string", DynamicStrings = new List<string> { "Yes", "No" }, DynamicInts = new List<int> { 2, 5 } });
            result.Add(new Models.BaseModel { DateTime = DateTime.Now, Double = 1.6, FirstKey = 2, Ignore = false, Int = 2, SecondKey = "Fourth", String = "string1", DynamicStrings = new List<string> { "soms", "Sometimes" } });
            result.Add(new Models.BaseModel { DateTime = DateTime.Now, Double = 1.7, FirstKey = 3, Ignore = true, Int = 3, SecondKey = "Eighth", String = "string2", DynamicStrings = new List<string> { "nooit", "Allways" } });
            result.Add(new Models.BaseModel { DateTime = DateTime.Now, Double = 1.5, FirstKey = 1, Ignore = false, Int = 1, SecondKey = "second", String = "string", DynamicStrings = new List<string> { "Yes", "No" } });
            result.Add(new Models.BaseModel { DateTime = DateTime.Now, Double = 1.6, FirstKey = 2, Ignore = false, Int = 2, SecondKey = "Fourth", String = "string1", DynamicStrings = new List<string> { "soms", "Sometimes" } });
            result.Add(new Models.BaseModel { DateTime = DateTime.Now, Double = 1.7, FirstKey = 3, Ignore = true, Int = 3, SecondKey = "Eighth", String = "string2", DynamicStrings = new List<string> { "nooit", "Allways" } });
            result.Add(new Models.BaseModel { DateTime = DateTime.Now, Double = 1.5, FirstKey = 1, Ignore = false, Int = 1, SecondKey = "second", String = "string", DynamicStrings = new List<string> { "Yes", "No" } });
            result.Add(new Models.BaseModel { DateTime = DateTime.Now, Double = 1.6, FirstKey = 2, Ignore = false, Int = 2, SecondKey = "Fourth", String = "string1", DynamicStrings = new List<string> { "soms", "Sometimes" } });
            result.Add(new Models.BaseModel { DateTime = DateTime.Now, Double = 1.7, FirstKey = 3, Ignore = true, Int = 3, SecondKey = "Eighth", String = "string2", DynamicStrings = new List<string> { "nooit", "Allways" } });
            result.Add(new Models.BaseModel { DateTime = DateTime.Now, Double = 1.5, FirstKey = 1, Ignore = false, Int = 1, SecondKey = "second", String = "string", DynamicStrings = new List<string> { "Yes", "No" } });
            result.Add(new Models.BaseModel { DateTime = DateTime.Now, Double = 1.6, FirstKey = 2, Ignore = false, Int = 2, SecondKey = "Fourth", String = "string1", DynamicStrings = new List<string> { "soms", "Sometimes" } });
            result.Add(new Models.BaseModel { DateTime = DateTime.Now, Double = 1.7, FirstKey = 3, Ignore = true, Int = 3, SecondKey = "Eighth", String = "string2", DynamicStrings = new List<string> { "nooit", "Allways" } });
            result.Add(new Models.BaseModel { DateTime = DateTime.Now, Double = 1.5, FirstKey = 1, Ignore = false, Int = 1, SecondKey = "second", String = "string", DynamicStrings = new List<string> { "Yes", "No" } });
            result.Add(new Models.BaseModel { DateTime = DateTime.Now, Double = 1.6, FirstKey = 2, Ignore = false, Int = 2, SecondKey = "Fourth", String = "string1", DynamicStrings = new List<string> { "soms", "Sometimes" } });
            result.Add(new Models.BaseModel { DateTime = DateTime.Now, Double = 1.7, FirstKey = 3, Ignore = true, Int = 3, SecondKey = "Eighth", String = "string2", DynamicStrings = new List<string> { "nooit", "Allways" } });
            result.Add(new Models.BaseModel { DateTime = DateTime.Now, Double = 1.5, FirstKey = 1, Ignore = false, Int = 1, SecondKey = "second", String = "string", DynamicStrings = new List<string> { "Yes", "No" } });
            result.Add(new Models.BaseModel { DateTime = DateTime.Now, Double = 1.6, FirstKey = 2, Ignore = false, Int = 2, SecondKey = "Fourth", String = "string1", DynamicStrings = new List<string> { "soms", "Sometimes" } });
            result.Add(new Models.BaseModel { DateTime = DateTime.Now, Double = 1.7, FirstKey = 3, Ignore = true, Int = 3, SecondKey = "Eighth", String = "string2", DynamicStrings = new List<string> { "nooit", "Allways" } });

            return result;
        }

        private void btn_LoadValidations_Click(object sender, RibbonControlEventArgs e)
        {
            var workbook = new ExcelWorkbook(Globals.ThisAddIn.Application.ActiveWorkbook);
            var worksheet = workbook.FindOrCreateWorksheet("Validations");
            var validValuesRange = worksheet.GetRange("A1").WriteValuesVertically(new List<string> { "Yes","No","Sometimes"});

            worksheet.GetRange("B1").Validation("=A1:A3");
            worksheet.GetRange("C1").Validation(validValuesRange);
        }

       





    }
}
