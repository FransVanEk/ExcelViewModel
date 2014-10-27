using Microsoft.Office.Interop.Excel;
using System;

namespace WannaApp.Excel.ExcelObjects
{
    public class ExcelConditionalFormatting : ExcelBaseObject, IDisposable
    {
        private FormatCondition _formatCondition;

        public ExcelConditionalFormatting(FormatCondition formatCondition)
        {
            this._formatCondition = formatCondition;
        }

        internal FormatCondition GetInteropVersion()
        {
            return this._formatCondition;
        }

        public FormatCondition interopVersion
        {
            get { return GetInteropVersion(); }
        }


        internal override void Dispose(bool disposing)
        {
            if (ExcelBaseObject.ReleaseComObjectOnDispose && GetInteropVersion() != null)
            {
                ReleaseComObject(GetInteropVersion());
            }
            if (disposing && _formatCondition != null)
            {
                _formatCondition = null;
            }

        }

    }
}
