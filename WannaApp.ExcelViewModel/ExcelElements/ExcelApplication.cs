using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WannaApp.ExcelViewModel.ExcelElements
{
    public class ExcelApplication
    {
        private _Application _excelApplication;

        public ExcelApplication(_Application excelApplication)
        {
            this._excelApplication = excelApplication;
        }

        internal _Application InteropApplication { get { return _excelApplication;  } }

 
    }
}
