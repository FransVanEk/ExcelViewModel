using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WannaApp.Excel.ExcelObjects
{
   public class ExcelWorksheet : ExcelBaseObject , IDisposable
    {
        private Worksheet _worksheet;

        public ExcelWorksheet(Worksheet worksheet)
        {
           
            this._worksheet = worksheet;
        }

        internal Worksheet GetInteropVersion()
        {
            return this._worksheet;
        }


        internal override void Dispose(bool disposing)
        {
            if (disposing  && ExcelBaseObject.ReleaseComObjectOnDispose && GetInteropVersion() != null)
            {
                ReleaseComObject(GetInteropVersion());
            }
            if (disposing && _worksheet != null)
            {
                _worksheet = null;
            }
            
        }

    }
}
