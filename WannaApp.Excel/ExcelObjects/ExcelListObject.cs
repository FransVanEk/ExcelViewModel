using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WannaApp.Excel.Extensions;

namespace WannaApp.Excel.ExcelObjects
{
    public class ExcelListObject : ExcelBaseObject, IDisposable
    {

        private ListObject _listObject;

        public ExcelListObject(ListObject listObject)
        {
            this._listObject = listObject;
        }

        internal ListObject GetInteropVersion()
        {
            return this._listObject;
        }

        internal override void Dispose(bool disposing)
        {
            if (ExcelBaseObject.ReleaseComObjectOnDispose && GetInteropVersion() != null)
            {
                ReleaseComObject(GetInteropVersion());
            }
            if (disposing && _listObject != null)
            {
                _listObject = null;
            }

        }

        public string[] Headers
        {
            get
            {
                return this.GetHeaderRange().GetHeaderValues(); 

            }
        }

        public object[,] dataValues
        {
            get
            {
                return this.GetDataRange().GetDataValues();
            }
        }
    }
}
