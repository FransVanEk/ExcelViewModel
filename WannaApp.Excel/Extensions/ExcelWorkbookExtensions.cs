using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WannaApp.Excel.ExcelObjects;

namespace WannaApp.Excel.Extensions
{
    public static class ExcelWorkbookExtensions
    {
        public static ExcelWorksheet AddNewWorksheet(this ExcelWorkbook workbook, string worksheetName)
        {
            ExcelWorksheet result = new ExcelWorksheet(workbook.GetInteropVersion().Worksheets.Add());
            result.Rename(worksheetName);
            return result;
        }

        public static ExcelWorksheet GetWorksheet(this ExcelWorkbook workbook, string worksheetName)
        {
            var worksheets = workbook.GetInteropVersion().Worksheets;
            ExcelWorksheet result = null;
            for (int i = 0; i < worksheets.Count; i++)
            {
                if (worksheets[i + 1].Name == worksheetName)
                {
                    result = new ExcelWorksheet(worksheets[i + 1]);
                }
            }
            return result;
        }

        public static ExcelWorksheet FindOrCreateWorksheet(this ExcelWorkbook workbook, string worksheetName)
        {
            ExcelWorksheet worksheet = workbook.GetWorksheet(worksheetName);
            if (worksheet == null) { worksheet = workbook.AddNewWorksheet(worksheetName);}
            return worksheet;
        }

        public static List<ExcelWorksheet> GetWorksheets(this ExcelWorkbook workbook)
        {
            var Worksheets = workbook.GetInteropVersion().Worksheets;
            var result = new List<ExcelWorksheet>();
            for (int i = 0; i < Worksheets.Count; i++)
            {
                result.Add(new ExcelWorksheet(Worksheets[i + 1])); //one based
            }

            return result;
        }

        public static ExcelWorkbook SaveAsWorkbook(this ExcelWorkbook workbook, string filename)
        {
            workbook.GetInteropVersion().SaveAs(filename);
            return workbook;
        }

        public static ExcelWorkbook SaveWorkbook(this ExcelWorkbook workbook)
        {
            workbook.GetInteropVersion().Save();
            return workbook;
        }

        public static ExcelListObject GetListObjectByName(this ExcelWorkbook workbook, string listname)
        {
            ExcelListObject result = null;
 

            workbook.GetWorksheets().ForEach(ws => {
                if (result == null && ws.ContainsListObjectByName(listname)) {
                    var list = ws.GetListObjectByName(listname);
                    if (list != null) { result = list; }
                }
        });

            return result;
        }

    }
}
