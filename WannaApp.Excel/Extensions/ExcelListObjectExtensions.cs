using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WannaApp.Excel.DataObjects;
using WannaApp.Excel.ExcelObjects;
using WannaApp.Excel.Extensions;


namespace WannaApp.Excel.Extensions
{
    public static class ExcelListObjectExtensions
    {
        const int xLengthIndex = 1;
        const int yLengthIndex = 0;

        public static ExcelListObject TableStyle(this ExcelListObject list, string styleName)
        {
            list.GetInteropVersion().TableStyle = styleName;
            return list;
        }

        public static ExcelListObjectHeaderRange GetHeaderRange(this ExcelListObject list)
        {
            return new ExcelListObjectHeaderRange(list.GetInteropVersion().HeaderRowRange);
        }

        public static ExcelListObjectDataRange GetDataRange(this ExcelListObject list)
        {
            var interopObject = list.GetInteropVersion();
            if (interopObject.DataBodyRange == null) {
                return new ExcelListObjectDataRange(interopObject.InsertRowRange);
            }
            else
            {
                return new ExcelListObjectDataRange(interopObject.DataBodyRange);
            }
            
        }

        public static IListObjectDataObject GetData(this ExcelListObject list)
        {
            var result =  new ListObjectDataObject();
            result.HeaderValues = list.GetHeaderRange().GetHeaderValues();
            result.DataValues = list.GetDataRange().GetDataValues();

            return result;

        }

        public static ExcelListObjectDataRange GetDataRangeFor(this ExcelListObject list, string columnName)
        {
            var index = list.GetColumnIndexFor(columnName) -1;
            if (index >= 0)
            {
                return list.GetDataRangeFor(index, index);
            }
            return null;
        }


          public static ExcelListObjectDataRange GetDataRangeFor(this ExcelListObject list, int startColumnIndex, int endColumnIndex)
        {
            return new ExcelListObjectDataRange(list.GetDataRange().GetColumnsFromRange(startColumnIndex, endColumnIndex).GetInteropVersion());
        }


        public static ExcelListObjectDataRange GetDataRangeFor(this ExcelListObject list, string startColumnName, string endColumnName)
        {
            var startindex = list.GetColumnIndexFor(startColumnName) -1;
            var endindex = list.GetColumnIndexFor(endColumnName) -1;
            if (startindex >= 0 && endindex >= 0)
            {
                return new ExcelListObjectDataRange(list.GetDataRange().GetColumnsFromRange(startindex, endindex).GetInteropVersion());
            }
            return null;
        }

        public static int GetColumnIndexFor(this ExcelListObject list, string columnName)
        {
            var headers = list.GetHeaderRange().ValuesAsArray();

            for (int x = 1; x <= headers.GetLength(xLengthIndex); x++)
            {
                for (int y = 1; y <= headers.GetLength(yLengthIndex); y++)
                {
                    if (headers[y, x].ToString().ToLower() == columnName.ToLower())
                    {
                        return x + 1;
                    }
                }
            }
            return -1;
        }
     
    }
}
