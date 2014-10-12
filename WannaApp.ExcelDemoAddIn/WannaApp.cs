using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using WannaApp.Excel.DataObjects;
using WannaApp.Excel.ExcelObjects;
using WannaApp.Excel.Extensions;
using System.Drawing;
using WannaApp.Excel.Helpers;

namespace WannaApp.ExcelDemoAddIn
{
    public partial class WannaApp
    {
        private void WannaApp_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void btn_LoadListObject_Click(object sender, RibbonControlEventArgs e)
        {
            var workbook = new ExcelWorkbook(Globals.ThisAddIn.Application.ActiveWorkbook);

            var list = workbook.FindOrCreateWorksheet("frans").CreateListObject("A1", GetData(), "eerste");
            var list2 = workbook.FindOrCreateWorksheet("frans").CreateListObject("G1", GetData(), "tweede");
            var list3 = workbook.FindOrCreateWorksheet("fransje").CreateListObject("A1", GetData(), "Derde");

            list.TableStyle("TableStyleLight12");
           
            
        }

        private ListObjectDataObject GetData()
        {
            var result = new ListObjectDataObject();
            result.HeaderValues = new string[] { "a", "b", "c", "d" };


            result.DataValues = new object[,] {
                                                  {"1","2","3","4"}
                                                , {"5","6","7","8"}
                                                , {"5","6","7","8"}
                                                , {"5","6","7","8"}
                                                , {"5","6","7","8"}
                                                , {"5","6","7","8"}
                                                , {"5","6","7","8"}
                                                , {"5","6","7","8"}
                                                , {"5","6","7","8"}
                                                , {"5","6","7","8"}
                                                , {"5","6","7","8"}
                                                , {"5","6","7","8"}
                                                , {"5","6","7","8"}
                                                , {"5","6","7","8"}
                                                , {"5","6","7","8"}
                                                , {"5","6","7","8"}
                                                , {"5","6","7","8"}
                                                , {"5","6","7","8"}
                                                , {"5","6","7","8"}
                                                , {"5","6","7","8"}
                                                , {"5","6","7","8"}
                                                , {"5","6","7","8"}
                                                , {"5","6","7","8"}
                                                , {"5","6","7","8"}
                                                , {"5","6","7","8"}
                                                , {"5","6","7","8"}
                                                , {"5","6","7","8"}
                                                , {"5","6","7","8"}
                                                , {"5","6","7","8"}
                                                , {"5","6","7","8"}
                                                , {"5","6","7","8"}
                                                , {"5","6","7","8"}
                                                , {"5","6","7","8"}
                                                , {"5","6","7","8"}
                                                , {"5","6","7","8"}
                                                , {"5","6","7","8"}
                                                , {"5","6","7","8"}
                                                , {"5","6","7","8"}
                                                , {"5","6","7","8"}
                                                , {"5","6","7","8"}
                                                , {"5","6","7","8"}
                                                , {"5","6","7","8"}
                                                , {"5","6","7","8"}
                                                , {"5","6","7","8"}
                                                , {"5","6","7","8"}
                                                , {"5","6","7","8"}
                                                , {"5","6","7","8"}
                                                , {"5","6","7","8"}
                                                , {"5","6","7","8"}
                                                , {"5","6","7","8"}
                                                , {"5","6","7","8"}
                                                , {"5","6","7","8"}
                                                , {"5","6","7","8"}
                                                , {"5","6","7","8"}
                                                , {"5","6","7","8"}
                                                , {"5","6","7","8"}
                                                , {"5","6","7","8"}
                                                , {"5","6","7","8"}
                                                , {"5","6","7","8"}
                                                , {"5","6","7","8"}
                                                , {"5","6","7","8"}
                                                , {"5","6","7","8"}
                                                , {"5","6","7","8"}
                                                , {"5","6","7","8"}
                                                , {"5","6","7","8"}
                                                , {"5","6","7","8"}
                                                , {"5","6","7","8"}
                                                , {"5","6","7","8"}
                                                , {"5","6","7","8"}
                                                , {"5","6","7","8"}
                                                , {"5","6","7","8"}
                                                , {"5","6","7","8"}
                                                , {"5","6","7","8"}
                                                , {"5","6","7","8"}
                                                , {"5","6","7","8"}
                                                , {"5","6","7","8"}
                                                , {"5","6","7","8"}
                                                };

           

            return result;
        }

        private void LoadObjects_Click(object sender, RibbonControlEventArgs e)
        {
           
            var workbook = new ExcelWorkbook(Globals.ThisAddIn.Application.ActiveWorkbook);
            var helper = GetHelperForBaseModel()
                .TransferToExcelFormat(GetTestDataNewTestObjects());

           var listobject =  workbook.FindOrCreateWorksheet("demo").CreateListObject("A1", helper, "eerste");
           listobject.GetHeaderRange().Orientation(45).Font(13);
           var test = listobject.GetHeaderRange().GetExcelRangeFor("frans", "Geert").Group().BackgroundColor(Color.LightBlue);

        }

      

        private IEnumerable<Models.BaseModel> GetTestDataNewTestObjects()
        {
            var result = new List<Models.BaseModel>();
            result.Add(new Models.BaseModel { DateTime = DateTime.Now, Double = 1.5, FirstKey = 1, Ignore = false, Int = 1, SecondKey = "tweede", String = "string", DynamicStrings = new List<string> { "ja", "nee" } , DynamicInts = new List<int>{2,5} });
            result.Add(new Models.BaseModel { DateTime = DateTime.Now, Double = 1.6, FirstKey = 2, Ignore = false, Int = 2, SecondKey = "vierde", String = "string1", DynamicStrings = new List<string> { "soms", "af en toe" } });
            result.Add(new Models.BaseModel { DateTime = DateTime.Now, Double = 1.7, FirstKey = 3, Ignore = true, Int = 3, SecondKey = "achtste", String = "string2", DynamicStrings = new List<string> { "nooit", "altijd" } });
            result.Add(new Models.BaseModel { DateTime = DateTime.Now, Double = 1.5, FirstKey = 1, Ignore = false, Int = 1, SecondKey = "tweede", String = "string", DynamicStrings = new List<string> { "ja", "nee" } });
            result.Add(new Models.BaseModel { DateTime = DateTime.Now, Double = 1.6, FirstKey = 2, Ignore = false, Int = 2, SecondKey = "vierde", String = "string1", DynamicStrings = new List<string> { "soms", "af en toe" } });
            result.Add(new Models.BaseModel { DateTime = DateTime.Now, Double = 1.7, FirstKey = 3, Ignore = true, Int = 3, SecondKey = "achtste", String = "string2", DynamicStrings = new List<string> { "nooit", "altijd" } });
            result.Add(new Models.BaseModel { DateTime = DateTime.Now, Double = 1.5, FirstKey = 1, Ignore = false, Int = 1, SecondKey = "tweede", String = "string", DynamicStrings = new List<string> { "ja", "nee" } });
            result.Add(new Models.BaseModel { DateTime = DateTime.Now, Double = 1.6, FirstKey = 2, Ignore = false, Int = 2, SecondKey = "vierde", String = "string1", DynamicStrings = new List<string> { "soms", "af en toe" } });
            result.Add(new Models.BaseModel { DateTime = DateTime.Now, Double = 1.7, FirstKey = 3, Ignore = true, Int = 3, SecondKey = "achtste", String = "string2", DynamicStrings = new List<string> { "nooit", "altijd" } });
            result.Add(new Models.BaseModel { DateTime = DateTime.Now, Double = 1.5, FirstKey = 1, Ignore = false, Int = 1, SecondKey = "tweede", String = "string", DynamicStrings = new List<string> { "ja", "nee" } });
            result.Add(new Models.BaseModel { DateTime = DateTime.Now, Double = 1.6, FirstKey = 2, Ignore = false, Int = 2, SecondKey = "vierde", String = "string1", DynamicStrings = new List<string> { "soms", "af en toe" } });
            result.Add(new Models.BaseModel { DateTime = DateTime.Now, Double = 1.7, FirstKey = 3, Ignore = true, Int = 3, SecondKey = "achtste", String = "string2", DynamicStrings = new List<string> { "nooit", "altijd" } });
            result.Add(new Models.BaseModel { DateTime = DateTime.Now, Double = 1.5, FirstKey = 1, Ignore = false, Int = 1, SecondKey = "tweede", String = "string", DynamicStrings = new List<string> { "ja", "nee" } });
            result.Add(new Models.BaseModel { DateTime = DateTime.Now, Double = 1.6, FirstKey = 2, Ignore = false, Int = 2, SecondKey = "vierde", String = "string1", DynamicStrings = new List<string> { "soms", "af en toe" } });
            result.Add(new Models.BaseModel { DateTime = DateTime.Now, Double = 1.7, FirstKey = 3, Ignore = true, Int = 3, SecondKey = "achtste", String = "string2", DynamicStrings = new List<string> { "nooit", "altijd" } });
            result.Add(new Models.BaseModel { DateTime = DateTime.Now, Double = 1.5, FirstKey = 1, Ignore = false, Int = 1, SecondKey = "tweede", String = "string", DynamicStrings = new List<string> { "ja", "nee" } });
            result.Add(new Models.BaseModel { DateTime = DateTime.Now, Double = 1.6, FirstKey = 2, Ignore = false, Int = 2, SecondKey = "vierde", String = "string1", DynamicStrings = new List<string> { "soms", "af en toe" } });
            result.Add(new Models.BaseModel { DateTime = DateTime.Now, Double = 1.7, FirstKey = 3, Ignore = true, Int = 3, SecondKey = "achtste", String = "string2", DynamicStrings = new List<string> { "nooit", "altijd" } });

            return result;
        }

        private void GetObjects_Click(object sender, RibbonControlEventArgs e)
        {
            var myExcelApp = new ExcelApplication(Globals.ThisAddIn.Application);
            var workbook = myExcelApp.FindWorkbook("Book1");
            var retrievedList = workbook.FindOrCreateWorksheet("demo").GetListObjectByName("eerste");
            var helper = GetHelperForBaseModel();
            var retrieveddata = helper.TransferFromExcelFormat(retrievedList).GetObjectsFromExcel();
            System.Windows.Forms.MessageBox.Show(string.Format("max value : {0}", GetMaxValueDynamicInts(retrieveddata)));
        }

        private int GetMaxValueDynamicInts(IEnumerable<Models.BaseModel> data)
        {
            return data.Where(d => d.DynamicInts != null && d.DynamicInts.Count > 0) 
                .Select(d => d.DynamicInts.Max()).Max();
        }

       

        private static TransferHelper<Models.BaseModel> GetHelperForBaseModel()
        {
            var helper = new Excel.Helpers.TransferHelper<Models.BaseModel>()
               .SetDynamicColumnsFor("DynamicStrings", new List<string> { "frans", "Geert" })
               .SetDynamicColumnsFor("DynamicInts", new List<string> { "Eerste", "Tweede" });
            return helper;
        }

        private void btn_loadDemo_Click(object sender, RibbonControlEventArgs e)
        {
            var workbook = new ExcelWorkbook(Globals.ThisAddIn.Application.ActiveWorkbook);
            var worksheet = workbook.FindOrCreateWorksheet("Demo HI");
            var listobject = worksheet.CreateListObject("A5", GetDemoData(), "MyDemoObjects");
            worksheet.GetRange("A1").BackgroundColor(Color.Red).SetValue("demo").Orientation(-10).Font(10);
            listobject.GetHeaderRange().AutoFit();
        }

        private TransferHelper<Models.Demo> GetDemoData()
        {
            return new TransferHelper<Models.Demo>().TransferToExcelFormat(GetDemoHIData());

        }

        private IEnumerable<Models.Demo> GetDemoHIData()
        {
            var result = new List<Models.Demo>();
            for (int i = 0; i < 100000; i++)
            {
                result.Add(new Models.Demo { Naam = "frans", sleutel = Guid.NewGuid().ToString()});
            }
            return result;
        }
    }
}
