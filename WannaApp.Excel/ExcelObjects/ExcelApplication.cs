using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WannaApp.Excel.ExcelObjects
{
    public class ExcelApplication  : ExcelBaseObject, IDisposable
    {
        private _Application _excel;

        public ExcelApplication(_Application excel)
        {
            this._excel = excel;
        }

        internal _Application GetInteropVersion()
        {
            return this._excel;
        }

        internal override void Dispose(bool disposing)
        {
            if (ExcelBaseObject.ReleaseComObjectOnDispose && GetInteropVersion() != null)
            {
                ReleaseComObject(GetInteropVersion());
            }
            if (disposing && _excel != null)
            {
                _excel = null;
            }

        }

       
    }
}
