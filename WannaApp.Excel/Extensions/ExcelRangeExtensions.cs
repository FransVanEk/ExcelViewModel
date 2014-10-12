using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Drawing;
using System.Threading.Tasks;
using WannaApp.Excel.DataObjects;
using WannaApp.Excel.ExcelObjects;

namespace WannaApp.Excel.Extensions
{
    public static class ExcelRangeExtensions
    {

        const int NumberOfHeaders = 1;
        const int xLengthIndex = 0;
        const int yLengthIndex = 1;

        public static ExcelListObject CreateListObject(this ExcelRange range, IListObjectDataObject data, string listObjectName)
        {
            var workingRange = range;
            if (isRangeSizeValid(workingRange,data) == false)
            {
               workingRange =  range.AdjustRangeSize(data);
            }
            workingRange.WriteData(data);
            return workingRange.ConvertIntoTable(listObjectName);
        }

        public static ExcelRange WriteData(this ExcelRange range,IListObjectDataObject data)
        {
            range.GetInteropVersion().Value2 = data.AllValues;
            return range;
        }

        public static ExcelListObject ConvertIntoTable(this ExcelRange range, string tableName)
        {
            Worksheet currentWorksheet = range.GetInteropVersion().Worksheet;
            var result = currentWorksheet.ListObjects.Add(XlListObjectSourceType.xlSrcRange,
                                            range.GetInteropVersion(), 
                                            System.Type.Missing, 
                                            XlYesNoGuess.xlYes, 
                                            System.Type.Missing);
            result.Name = tableName;

            return new ExcelListObject(result);

        }

        private static ExcelRange AdjustRangeSize(this ExcelRange range, IListObjectDataObject data)
        {
            Range topLeft = range.GetLeftTopCell().GetInteropVersion();
            var columnAbsoluteIndex = topLeft.Column;
            var rowAbsoluteIndex = topLeft.Row;
            var columnsWide = data.HeaderValues.Length;
            var rowsHeight = data.DataValues.GetLength(xLengthIndex);
            Range bottomRight = topLeft.Worksheet.Cells[rowAbsoluteIndex + rowsHeight + NumberOfHeaders -1, columnAbsoluteIndex + columnsWide - 1]; //  -1 reason one base
            return new ExcelRange(topLeft.Worksheet.get_Range(topLeft, bottomRight));
        }

        private static bool isRangeSizeValid(ExcelRange range, IListObjectDataObject data)
        {
            var internalRange = range.GetInteropVersion();

            return (internalRange.Columns.Count == data.HeaderValues.Length && internalRange.Rows.Count == data.DataValues.GetLength(yLengthIndex) + NumberOfHeaders); 
        }

        private static ExcelRange GetBottomRightCell(this ExcelRange range)
        {
            var internalRange = range.GetInteropVersion();

            return new ExcelRange(internalRange.Cells[internalRange.Rows.Count, internalRange.Columns.Count]);
        }

        public static ExcelRange GetLeftTopCell(this ExcelRange range)
        {
             return new ExcelRange(range.GetInteropVersion().Cells[1,1]);
        }

        public static object[,] ValuesAsArray(this ExcelRange range)
        {
            return range.GetInteropVersion().Value2;
        }
        
        public static ExcelRange Merge(this ExcelRange range)
        {
            range.GetInteropVersion().Merge(false);
            return range;
        }

        public static ExcelRange Orientation(this ExcelRange range, int angle)
        {
            range.GetInteropVersion().Orientation = angle;
            return range;
        }

        public static ExcelRange FontStyling(this ExcelRange range,Microsoft.Office.Interop.Excel.Font fontStyle)
        {
            range.GetInteropVersion().Font.FontStyle = fontStyle;
            return range;
        }

        public static ExcelRange Font(this ExcelRange range, int fontSize)
        {
            range.GetInteropVersion().Font.Size = fontSize;
            return range;
        }

        public static ExcelRange Wrap(this ExcelRange range, bool wrapText)
        {
            range.GetInteropVersion().WrapText = wrapText;
            return range;
        }

        public static ExcelRange AutoFit(this ExcelRange range)
        {
            range.GetInteropVersion().EntireColumn.AutoFit();
            return range;
        }

        public static ExcelRange Group(this ExcelRange range)
        {
            range.GetInteropVersion().Group();
            return range;
        }

        public static ExcelRange BackgroundColor(this ExcelRange range, Color color)
        {
            range.GetInteropVersion().Interior.Color = color;
            return range;
        }

        public static string Address(this ExcelRange range)
        {
            return range.GetInteropVersion().get_Address();
        }

        public static ExcelRange SetValue(this ExcelRange range, object value)
        {
            range.GetInteropVersion().Value2 = value;
            return range;
        }

        public static ExcelRange Format(this ExcelRange range, string format)
        {
            range.GetInteropVersion().NumberFormat = format;
            return range;
        }
    }
}
