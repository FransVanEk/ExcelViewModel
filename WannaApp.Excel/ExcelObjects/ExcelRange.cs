using Microsoft.Office.Interop.Excel;
using System;

namespace WannaApp.Excel.ExcelObjects
{
    public class ExcelRange : ExcelBaseObject , IDisposable
    {
        private Range _range;

        public ExcelRange(Range range)
        {
            this._range = range;
        }

        internal Range GetInteropVersion()
        {
            return this._range;
        }



        internal override void Dispose(bool disposing)
        {
            if (ExcelBaseObject.ReleaseComObjectOnDispose && GetInteropVersion() != null)
            {
                ReleaseComObject(GetInteropVersion());
            }
            if (disposing && _range != null)
            {
                _range = null;
            }

        }

    }
}
