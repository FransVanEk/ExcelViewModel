using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WannaApp.Excel.ExcelObjects
{
    public class ExcelWorkbook : ExcelBaseObject, IDisposable
    {
        private Workbook _workbook;

        public ExcelWorkbook(Workbook workbook)
        {
            this._workbook = workbook;
        }

        internal Workbook GetInteropVersion()
        {
            return this._workbook;
        }

        internal override void Dispose(bool disposing)
        {
            if (disposing && ExcelBaseObject.ReleaseComObjectOnDispose && GetInteropVersion() != null)
            {
                ReleaseComObject(GetInteropVersion());
            }
            if (disposing && _workbook != null)
            {
                _workbook = null;
            }

        }
    }
}
