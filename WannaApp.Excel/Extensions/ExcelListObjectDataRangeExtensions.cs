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
    public static class ExcelListObjectDataRangeExtensions
    {
        public static object[,] GetDataValues(this ExcelListObjectDataRange range)
        {
            return range.GetInteropVersion().Value2;
        }

    }
}
