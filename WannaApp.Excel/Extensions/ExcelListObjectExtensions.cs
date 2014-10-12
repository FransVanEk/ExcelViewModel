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
            return new ExcelListObjectDataRange(list.GetInteropVersion().DataBodyRange);
        }

        public static IListObjectDataObject GetData(this ExcelListObject list)
        {
            var result =  new ListObjectDataObject();
            result.HeaderValues = list.GetHeaderRange().GetHeaderValues();
            result.DataValues = list.GetDataRange().GetDataValues();

            return result;

        }

     
    }
}
