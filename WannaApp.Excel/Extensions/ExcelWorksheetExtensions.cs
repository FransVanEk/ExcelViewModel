using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WannaApp.Excel.DataObjects;
using WannaApp.Excel.ExcelObjects;


namespace WannaApp.Excel.Extensions
{
    public static class ExcelWorksheetExtensions
    {
        public static ExcelWorksheet Rename(this ExcelWorksheet worksheet,string newName)
        {
            worksheet.GetInteropVersion().Name = newName;
            return worksheet;
        }

        public static void Delete(this ExcelWorksheet worksheet)
        {
            worksheet.GetInteropVersion().Delete();
        }

        public static ExcelWorksheet PlaceBefore(this ExcelWorksheet worksheet, ExcelWorksheet beforeWorksheet)
        {
            worksheet.GetInteropVersion().Move(beforeWorksheet.GetInteropVersion(),Type.Missing);
            return worksheet;
        }

        public static ExcelWorksheet PlaceAfter(this ExcelWorksheet worksheet, ExcelWorksheet afterWorksheet)
        {
            worksheet.GetInteropVersion().Move(afterWorksheet.GetInteropVersion(), Type.Missing);
            return worksheet;
        }

        public static string Name(this ExcelWorksheet worksheet)
        {
            return worksheet.GetInteropVersion().Name;
        }

        public static ExcelListObject CreateListObject(this ExcelWorksheet worksheet, ExcelRange leftTopTargetCell, IListObjectDataObject data, string listObjectName)
        {
            return leftTopTargetCell.CreateListObject(data,listObjectName);
        }

        public static ExcelListObject CreateListObject(this ExcelWorksheet worksheet, string leftTopTargetCellAddress, IListObjectDataObject data, string listObjectName)
        {
            return worksheet.CreateListObject(worksheet.GetRange(leftTopTargetCellAddress), data,listObjectName);
        }

        public static ExcelRange GetRange(this ExcelWorksheet worksheet, string Address)
        {
            return new ExcelRange(worksheet.GetInteropVersion().Range[Address]);
        }

        public static ExcelListObject GetListObjectByName(this ExcelWorksheet worksheet, string ListObjectName)
        {
            return new ExcelListObject(worksheet.GetInteropVersion().ListObjects[ListObjectName]);
        }
    }
}
