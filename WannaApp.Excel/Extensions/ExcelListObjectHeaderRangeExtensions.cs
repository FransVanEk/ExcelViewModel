using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WannaApp.Excel.DataObjects;
using WannaApp.Excel.ExcelObjects;

namespace WannaApp.Excel.Extensions
{
    public static class ExcelListObjectHeaderRangeExtensions
    {

        public static string[] GetHeaderValues(this ExcelListObjectHeaderRange range)
        {
            object[,] allHeaderValues = range.GetInteropVersion().Value2;
            List<string> result = new List<string>();
            for (int i = 1; i <= allHeaderValues.GetLength(1); i++)
            {
                result.Add(allHeaderValues[1, i].ToString());
            }
            return result.ToArray();
        }

        public static int GetIndexForColumn(this ExcelListObjectHeaderRange headerRange, string columnName)
        {
            return headerRange.GetHeaderValues().ToList().IndexOf(columnName);
        }

        public static ExcelRange GetExcelRangeFor(this ExcelListObjectHeaderRange range, string startColumn, string endColumn)
        {
            var baseColumnIndex = range.GetInteropVersion().Column;
            var baseRowIndex = range.GetInteropVersion().Row;
            var startColumnIndex = range.GetIndexForColumn(startColumn) + baseColumnIndex;
            var endColumnIndex = range.GetIndexForColumn(endColumn) + baseColumnIndex;
            Worksheet worksheet = range.GetInteropVersion().Worksheet;
            Range startCell = worksheet.Cells[baseRowIndex, startColumnIndex];
            Range endCell = worksheet.Cells[baseRowIndex,endColumnIndex];
            return new ExcelRange((Range)worksheet.get_Range(startCell, endCell));

        }
    }
}
