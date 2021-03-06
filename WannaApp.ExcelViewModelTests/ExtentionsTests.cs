﻿using System;
using iExcel = Microsoft.Office.Interop.Excel;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Microsoft.Office.Interop.Excel;
using System.Linq;
using WannaApp.Excel.ExcelObjects;
using WannaApp.Excel.Extensions;
using WannaApp.Excel.DataObjects;
using WannaApp.Excel.Helpers;
using System.Collections.Generic;
using WannaApp.Excel.Helpers.MappingHelpers;
using System.IO;


namespace WannaApp.ExcelViewModelTests
{
    [TestClass]
    public class ExtentionsTests
    {
        private string _filename = @"c:\temp\test.xlsx";
        public Application StartExcel()
        {
            iExcel.Application instance = null;
            try
            {
                instance = (iExcel.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");
            }
            catch (System.Runtime.InteropServices.COMException)
            {
                instance = new iExcel.Application();
            }

            return instance;
        }

        [TestMethod]
        public void GetWorkBooksTest()
        {
            var excel = StartExcel();
            var test = new ExcelApplication(excel);
            var workbooks = test.GetWorkbooks();
           
            Assert.AreNotEqual(null, workbooks);
            excel.Quit();
        }

        [TestMethod]
        public void CreateWorkbookTest()
        {
            var excel = StartExcel();
            var test = new ExcelApplication(excel);
            var NumberOfworkbooks = test.GetWorkbooks().Count;
            var workbook = test.AddNewWorkbook();
            Assert.AreNotEqual(null, workbook);
            Assert.AreNotEqual(NumberOfworkbooks + 1, test.GetWorkbooks());
            workbook.Close(false);
            excel.Quit();
        }

        [TestMethod]
        public void GetWorkbookByNameTest()
        {
            var excel = StartExcel();
            var test = new ExcelApplication(excel);
            var workbook = test.GetLoadedWorkbook("Demo.xls");
            excel.Quit();
            Assert.AreEqual(null, workbook);
        }

        [TestMethod]
        public void WorkBookWorkSheetManipulationsTest()
        {
            RemoveFile();
            var excel = StartExcel();
            var test = new ExcelApplication(excel);
            var workbook = test.AddNewWorkbook();
            var numberOfSheets = workbook.GetWorksheets().Count;
            var newsheet = workbook.AddNewWorksheet("Test");
            var retrieveNewSheet = workbook.GetWorksheet("Test");
            workbook.SaveAsWorkbook(_filename);
            Assert.AreEqual(numberOfSheets + 1, workbook.GetWorksheets().Count);
            Assert.AreNotEqual(null, newsheet);
            Assert.AreNotEqual(null, retrieveNewSheet);
            workbook.Close(false);
            excel.Quit();
        }

        [TestMethod]
        public void WorkBookWorkSheetCreateListObject()
        {
            RemoveFile();
            var excel = StartExcel();
            var test = new ExcelApplication(excel);
            var workbook = test.AddNewWorkbook();
            var list = workbook.FindOrCreateWorksheet("frans").CreateListObject("A1", GetData(), "eerste");
            var list2 = workbook.FindOrCreateWorksheet("frans").CreateListObject("G1", GetData(), "tweede");
            var list3 = workbook.FindOrCreateWorksheet("fransje").CreateListObject("A1", GetData(), "Derde");
            var newsheet = workbook.GetWorksheet("fransje");
            workbook.SaveAsWorkbook(_filename);
            excel.Quit();
            Assert.AreNotEqual(null, newsheet);
            Assert.AreNotEqual(null, list);
            
        }

        private void RemoveFile()
        {
            new FileInfo(_filename).Delete();
        }

        [TestMethod]
        public void MappingManagerTest()
        {
            var manager = new WannaApp.Excel.Helpers.MappingHelpers.TransferHelper<TestObjecten.NewTestObject>().SetDynamicColumnsFor("DynamicStrings", new List<string> { "frans", "Geert" });
            Assert.AreNotEqual(null, manager.HeaderValues);
        }

        [TestMethod]
        public void MappingManagerTestWithData()
        {
            var manager = new WannaApp.Excel.Helpers.MappingHelpers.TransferHelper<TestObjecten.NewTestObject>().SetDynamicColumnsFor("DynamicStrings", new List<string> { "frans", "Geert" });
            manager.TransferToExcelFormat(GetTestDataNewTestObjects());
            

            Assert.AreNotEqual(null, manager.HeaderValues);
            Assert.AreNotEqual(null, manager.DataValues);
            Assert.AreNotEqual(null, manager.AllValues);
        }


        [TestMethod]
        public void WorkBookWorkSheetCreateListObjectFromObjects()
        {
            RemoveFile();
            var excel = StartExcel();
            var test = new ExcelApplication(excel);
            var workbook = test.AddNewWorkbook();
            var helper = new WannaApp.Excel.Helpers.MappingHelpers.TransferHelper<TestObjecten.NewTestObject>()
                .SetDynamicColumnsFor("DynamicStrings", new List<string> { "frans", "Geert" })
                .TransferToExcelFormat(GetTestDataNewTestObjects());
   
            var list = workbook.FindOrCreateWorksheet("frans").CreateListObject("A1", helper, "eerste");
            var retrievedList =  workbook.GetListObjectByName("eerste");

            var retrieveddata =  helper.TransferFromExcelFormat(retrievedList).GetObjectsFromExcel();
            workbook.SaveAsWorkbook(_filename);
            excel.Quit();
           
            Assert.AreNotEqual(null, list);

        }

        [TestMethod]
        public void WorkBookWorkSheetContainesListObject()
        {
            var excel = StartExcel();
            var test = new ExcelApplication(excel);
            var workbook = test.AddNewWorkbook();
            var helper = new WannaApp.Excel.Helpers.MappingHelpers.TransferHelper<TestObjecten.NewTestObject>()
                .SetDynamicColumnsFor("DynamicStrings", new List<string> { "frans", "Geert" })
                .TransferToExcelFormat(GetTestDataNewTestObjects());
            var worksheet = workbook.FindOrCreateWorksheet("frans");
            var list = worksheet.CreateListObject("A1", helper, "ListObjectName");
            var found = worksheet.ContainsListObjectByName("ListObjectName");
            var notFound = worksheet.ContainsListObjectByName("noExistingListName");
            var workbookFound = workbook.GetListObjectByName("ListObjectName");
            var workbookNotFound =   workbook.GetListObjectByName("noExistingListName");

            Assert.AreEqual(true, found);
            Assert.AreEqual(false,notFound);
            Assert.AreNotEqual(null, workbookFound);
            Assert.AreEqual(null, workbookNotFound);
            workbook.Close(false);
            excel.Quit();
        }

        [TestMethod]
        public void HeaderMappinginfoTest()
        {
            var helper = new ExcelToObjectMappingHelper();
            var transferHelper = new WannaApp.Excel.Helpers.MappingHelpers.TransferHelper<TestObjecten.NewTestObject>()
                .SetDynamicColumnsFor("DynamicStrings", new List<string> { "frans", "Geert" });
            var result = helper.Process(new string[] { "Eerste Sleutel", "Tweede sleutel", "Tekst", "Getal", "Datum", "Decimaal", "frans", "Geert" ,"DynamicInts" }, transferHelper.MappingInfoToExcel);
          
            Assert.AreNotEqual(null, result);
            Assert.AreNotEqual(0, result.Count);
        }

 



        private IEnumerable<TestObjecten.NewTestObject> GetTestDataNewTestObjects()
        {
            var result = new List<TestObjecten.NewTestObject>();
            result.Add(new TestObjecten.NewTestObject { DateTime = DateTime.Now, Double = 1.5, FirstKey = 1, Ignore = false, Int = 1, SecondKey = "tweede", String = "string", DynamicStrings = new List<string> { "ja", "nee" }, Guid = Guid.NewGuid() });
            result.Add(new TestObjecten.NewTestObject { DateTime = DateTime.Now, Double = 1.6, FirstKey = 2, Ignore = false, Int = 2, SecondKey = "vierde", String = "string1", DynamicStrings = new List<string> { "soms", "af en toe" }, Guid = Guid.NewGuid() });
            result.Add(new TestObjecten.NewTestObject { DateTime = DateTime.Now, Double = 1.7, FirstKey = 3, Ignore = true, Int = 3, SecondKey = "achtste", String = "string2", DynamicStrings = new List<string> { "nooit", "altijd" }, Guid = Guid.NewGuid() });
            result.Add(new TestObjecten.NewTestObject { DateTime = DateTime.Now, Double = 1.5, FirstKey = 1, Ignore = false, Int = 1, SecondKey = "tweede", String = "string", DynamicStrings = new List<string> { "ja", "nee" }, Guid = Guid.NewGuid() });
            result.Add(new TestObjecten.NewTestObject { DateTime = DateTime.Now, Double = 1.6, FirstKey = 2, Ignore = false, Int = 2, SecondKey = "vierde", String = "string1", DynamicStrings = new List<string> { "soms", "af en toe" }, Guid = Guid.NewGuid() });
            result.Add(new TestObjecten.NewTestObject { DateTime = DateTime.Now, Double = 1.7, FirstKey = 3, Ignore = true, Int = 3, SecondKey = "achtste", String = "string2", DynamicStrings = new List<string> { "nooit", "altijd" }, Guid = Guid.NewGuid() });
            result.Add(new TestObjecten.NewTestObject { DateTime = DateTime.Now, Double = 1.5, FirstKey = 1, Ignore = false, Int = 1, SecondKey = "tweede", String = "string", DynamicStrings = new List<string> { "ja", "nee" }, Guid = Guid.NewGuid() });
            result.Add(new TestObjecten.NewTestObject { DateTime = DateTime.Now, Double = 1.6, FirstKey = 2, Ignore = false, Int = 2, SecondKey = "vierde", String = "string1", DynamicStrings = new List<string> { "soms", "af en toe" }, Guid = Guid.NewGuid() });
            result.Add(new TestObjecten.NewTestObject { DateTime = DateTime.Now, Double = 1.7, FirstKey = 3, Ignore = true, Int = 3, SecondKey = "achtste", String = "string2", DynamicStrings = new List<string> { "nooit", "altijd" }, Guid = Guid.NewGuid() });
            result.Add(new TestObjecten.NewTestObject { DateTime = DateTime.Now, Double = 1.5, FirstKey = 1, Ignore = false, Int = 1, SecondKey = "tweede", String = "string", DynamicStrings = new List<string> { "ja", "nee" }, Guid = Guid.NewGuid() });
            result.Add(new TestObjecten.NewTestObject { DateTime = DateTime.Now, Double = 1.6, FirstKey = 2, Ignore = false, Int = 2, SecondKey = "vierde", String = "string1", DynamicStrings = new List<string> { "soms", "af en toe" }, Guid = Guid.NewGuid() });
            result.Add(new TestObjecten.NewTestObject { DateTime = DateTime.Now, Double = 1.7, FirstKey = 3, Ignore = true, Int = 3, SecondKey = "achtste", String = "string2", DynamicStrings = new List<string> { "nooit", "altijd" }, Guid = Guid.NewGuid() });
            result.Add(new TestObjecten.NewTestObject { DateTime = DateTime.Now, Double = 1.5, FirstKey = 1, Ignore = false, Int = 1, SecondKey = "tweede", String = "string", DynamicStrings = new List<string> { "ja", "nee" }, Guid = Guid.NewGuid() });
            result.Add(new TestObjecten.NewTestObject { DateTime = DateTime.Now, Double = 1.6, FirstKey = 2, Ignore = false, Int = 2, SecondKey = "vierde", String = "string1", DynamicStrings = new List<string> { "soms", "af en toe" }, Guid = Guid.NewGuid() });
            result.Add(new TestObjecten.NewTestObject { DateTime = DateTime.Now, Double = 1.7, FirstKey = 3, Ignore = true, Int = 3, SecondKey = "achtste", String = "string2", DynamicStrings = new List<string> { "nooit", "altijd" }, Guid = Guid.NewGuid() });
            result.Add(new TestObjecten.NewTestObject { DateTime = DateTime.Now, Double = 1.5, FirstKey = 1, Ignore = false, Int = 1, SecondKey = "tweede", String = "string", DynamicStrings = new List<string> { "ja", "nee" }, Guid = Guid.NewGuid() });
            result.Add(new TestObjecten.NewTestObject { DateTime = DateTime.Now, Double = 1.6, FirstKey = 2, Ignore = false, Int = 2, SecondKey = "vierde", String = "string1", DynamicStrings = new List<string> { "soms", "af en toe" }, Guid = Guid.NewGuid() });
            result.Add(new TestObjecten.NewTestObject { DateTime = DateTime.Now, Double = 1.7, FirstKey = 3, Ignore = true, Int = 3, SecondKey = "achtste", String = "string2", DynamicStrings = new List<string> { "nooit", "altijd" }, Guid = Guid.NewGuid() });

            return result;
        }


        private ListObjectDataObject GetData()
        {
            var result = new ListObjectDataObject();
            result.HeaderValues = new string[] { "a", "b", "c", "d" };

  
            result.DataValues = new string[,] {
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

    }
}
