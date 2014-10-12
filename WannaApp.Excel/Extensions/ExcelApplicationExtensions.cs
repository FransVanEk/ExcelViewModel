using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WannaApp.Excel.ExcelObjects;
using WannaApp.Excel.Extensions;

namespace WannaApp.Excel.Extensions
{
    public static class ExcelApplicationExtensions
    {
        public static List<ExcelWorkbook>  GetWorkbooks(this ExcelApplication application)
        {
            var WorkBooks = application.GetInteropVersion().Workbooks ;
            var result = new List<ExcelWorkbook>();
            for (int i = 0; i < WorkBooks.Count; i++)
            {
                result.Add(new ExcelWorkbook(WorkBooks[i+1])); //one based
            }

            return result;
        }

        public static ExcelWorkbook GetLoadedWorkbook(this ExcelApplication application,string filename)
        {
            var WorkBooks = application.GetInteropVersion().Workbooks ;
            ExcelWorkbook result = null;
            for (int i = 0; i < WorkBooks.Count; i++)
            {
                if(WorkBooks[i+1].Name == filename ) //one based
                {
                    result = new ExcelWorkbook(WorkBooks[i+1]);
                    break;
                }
            }

            return result;
        }

        public static ExcelWorkbook AddNewWorkbook(this ExcelApplication application)
        {
            return new ExcelWorkbook(  application.GetInteropVersion().Workbooks.Add());
        }

        public static ExcelWorkbook OpenWorkbook(this ExcelApplication application, string filename)
        {
            return new ExcelWorkbook(application.GetInteropVersion().Workbooks.Open(filename));
        }

        public static List<ExcelWorksheet> GetWorksheets(this ExcelApplication application)
        {
            var Worksheets = application.GetInteropVersion().Worksheets;
            var result = new List<ExcelWorksheet>();
            for (int i = 0; i < Worksheets.Count; i++)
            {
                result.Add(new ExcelWorksheet(Worksheets[i + 1])); //one based
            }

            return result;
        }

        public static ExcelWorkbook FindWorkbook(this ExcelApplication application, string name)
        {
            var workbooks =  application.GetInteropVersion().Workbooks;
            ExcelWorkbook result = null;
            for (int i = 1; i <= workbooks.Count; i++)
            {
                if (workbooks[i].Name == name)
                {
                    result = new ExcelWorkbook(workbooks[i]);
                    break;
                }
            }
            return result;
        }

        public static ExcelWorksheet GetWorksheet(this ExcelApplication application, string worksheetName)
        {
            var Worksheets = application.GetInteropVersion().Worksheets;
            ExcelWorksheet result = null;
            for (int i = 1; i <= Worksheets.Count; i++)
            {
                if (Worksheets[i].Name == worksheetName) //one based
                {
                    result = new ExcelWorksheet(Worksheets[i]);
                    break;
                }
            }

            return result;
        }

        public static void Close(this ExcelApplication application)
        {
            application.GetInteropVersion().Quit();
        }
    }
}
