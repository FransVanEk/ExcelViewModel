using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Drawing;
using System.Threading.Tasks;
using WannaApp.Excel.DataObjects;
using WannaApp.Excel.ExcelObjects;
using System.Runtime.Remoting;

namespace WannaApp.Excel.Extensions
{
    public static class ExcelRangeExtensions
    {

        const int NumberOfHeaders = 1;
        const int xLengthIndex = 1;
        const int yLengthIndex = 0;

        public static ExcelListObject CreateListObject(this ExcelRange range, IListObjectDataObject data, string listObjectName)
        {
            var workingRange = range.WriteData(data.AllValues);
            return workingRange.ConvertIntoTable(listObjectName);
        }

        public static ExcelRange WriteData(this ExcelRange range, object[,] data)
        {
            var result = range;
            if (isRangeSizeValid(range, data) == false)
            {
                result = range.GetRangeForSize(data.GetLength(yLengthIndex), data.GetLength(xLengthIndex));
            }
            result.SetValue(data);
            return result;
        }

        public static ExcelRange GetRangeForSize(this ExcelRange range, int numberOrRows, int numberOfColumns)
        {
            Range topLeft = range.GetLeftTopCell().GetInteropVersion();
            return range.ExtendRangeSize(numberOrRows - 1, numberOfColumns - 1);
        }

        public static ExcelRange WriteData(this ExcelRange range, IListObjectDataObject data)
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
            var rowsHeight = data.DataValues.GetLength(yLengthIndex);
            Range bottomRight = topLeft.Worksheet.Cells[rowAbsoluteIndex + rowsHeight + NumberOfHeaders - 1, columnAbsoluteIndex + columnsWide - 1]; //  -1 reason one base
            return new ExcelRange(topLeft.Worksheet.get_Range(topLeft, bottomRight));
        }

        private static ExcelRange ExtendRangeSize(this ExcelRange range, int extendRows, int extendColumns)
        {
            var rangeInteropVersion = range.GetInteropVersion();
            Range topLeft = range.GetLeftTopCell().GetInteropVersion();
            var columnAbsoluteIndex = topLeft.Column;
            var rowAbsoluteIndex = topLeft.Row;
            var columnsWide = rangeInteropVersion.Columns.Count + extendColumns;
            var rowsHeight = rangeInteropVersion.Rows.Count + extendRows;
            Range bottomRight = topLeft.Worksheet.Cells[rowAbsoluteIndex + rowsHeight - 1, columnAbsoluteIndex + columnsWide - 1]; //  -1 reason one base
            return new ExcelRange(topLeft.Worksheet.get_Range(topLeft, bottomRight));
        }

        private static bool isRangeSizeValid(ExcelRange range, IListObjectDataObject data)
        {
            return isRangeSizeValid(range, data.AllValues);
        }

        private static bool isRangeSizeValid(ExcelRange range, object[,] data)
        {
            var internalRange = range.GetInteropVersion();

            return (internalRange.Columns.Count == data.GetLength(xLengthIndex) && internalRange.Rows.Count == data.GetLength(yLengthIndex) + NumberOfHeaders);
        }

        private static ExcelRange GetBottomRightCell(this ExcelRange range)
        {
            var internalRange = range.GetInteropVersion();

            return new ExcelRange(internalRange.Cells[internalRange.Rows.Count, internalRange.Columns.Count]);
        }

        public static ExcelRange GetLeftTopCell(this ExcelRange range)
        {
            return new ExcelRange(range.GetInteropVersion().Cells[1, 1]);
        }

        public static object[,] ValuesAsArray(this ExcelRange range)
        {
            if (range.GetInteropVersion().Cells.Count == 1)
            {
                var result = new Object[2, 2];
                result[1, 1] = range.GetInteropVersion().Value2;
                return result;
            }
            else
            {
                return range.GetInteropVersion().Value2;
            }

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

        public static ExcelRange FontStyling(this ExcelRange range, Microsoft.Office.Interop.Excel.Font fontStyle)
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

        public static ExcelRange Width(this ExcelRange range, int widthInPoints)
        {
            range.GetInteropVersion().ColumnWidth = widthInPoints;
            return range;
        }

        public static ExcelRange Validation(this ExcelRange range,
                               string validationFormula,
                               string errorTitle,
                               string errorText
                                        )
        {
            var interopRange = range.GetInteropVersion();

            interopRange.Validation.Delete();
            interopRange.Validation.Add(
                XlDVType.xlValidateList,
                XlDVAlertStyle.xlValidAlertInformation,
                XlFormatConditionOperator.xlBetween,
                validationFormula,
                Type.Missing);

            interopRange.Validation.IgnoreBlank = true;
            interopRange.Validation.ErrorMessage = string.IsNullOrEmpty(errorText) ? string.Empty : errorText;
            interopRange.Validation.ErrorTitle = string.IsNullOrEmpty(errorTitle) ? "Error" : errorTitle;
            interopRange.Validation.InCellDropdown = true;
            interopRange.Validation.ShowError = string.IsNullOrEmpty(errorText);

            return range;
        }

        public static ExcelRange Validation(this ExcelRange range,
                          string validationFormula
                                   )
        {
            return range.Validation(validationFormula, string.Empty, string.Empty);

        }

        public static ExcelRange Validation(this ExcelRange range,
                           ExcelRange validValues)
        {
            return range.Validation(validValues, string.Empty, string.Empty);
        }

        public static ExcelRange Validation(this ExcelRange range,
                              ExcelRange validValues,
                              string errorTitle,
                              string errorText
                                       )
        {
            return range.Validation(String.Format("='{0}'!{1}",
                  validValues.GetInteropVersion().Worksheet.Name,
                  validValues.GetInteropVersion().get_AddressLocal()),
                  errorTitle,
                  errorText);
        }

        public static ExcelRange WriteValuesVertically(this ExcelRange range, List<string> values)
        {
            var data = new object[values.Count, 1];
            var extendedRange = range.ExtendRangeSize(values.Count - 1, 0);
            values.ForEach(v => data[values.IndexOf(v), 0] = v);
            return extendedRange.SetValue(data);
        }

        public static ExcelRange WriteValuesHorizontally(this ExcelRange range, List<string> values)
        {
            var data = new object[1, values.Count];
            var extendedRange = range.ExtendRangeSize(0, values.Count - 1);
            values.ForEach(v => data[0, values.IndexOf(v)] = v);
            return extendedRange.SetValue(data);
        }

        public static ExcelRange GetColumnsFromRange(this ExcelRange range, int startcolumn, int endColumn)
        {
            var startrange = range.GetLeftTopCell().GetSingleCellByOffset(0, startcolumn - 1);
            return startrange.ExtendRangeSize(range.GetInteropVersion().Rows.Count - 1, endColumn - startcolumn);
        }


        public static ExcelRange GetSingleCellByOffset(this ExcelRange range, int rowOffset, int columnOffset)
        {
            var rangeInteropVersion = range.GetInteropVersion();
            var currentRow = rangeInteropVersion.Row;
            var currentColumn = rangeInteropVersion.Column;
            return new ExcelRange(rangeInteropVersion.Worksheet.Cells[currentRow + rowOffset, currentColumn + columnOffset]);
        }

    }
}
