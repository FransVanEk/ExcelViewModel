using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WannaApp.ExcelViewModel.Interfaces;

namespace WannaApp.ExcelViewModel.ExcelElements
{
    public class ExcelWorksheet
    {
        private Worksheet _interopWorksheet;
        private IExcelManager _excelManager;
        private ListObjectManager _listObjectManager;

        public ExcelWorksheet(IExcelManager excelManager, Worksheet worksheet)
        {
            this._interopWorksheet = worksheet;
            this._excelManager = excelManager;
        }

        public Worksheet InteropWorkSheet
        {
            get
            {
                return _interopWorksheet;
            }
        }

        public string Name
        {
            get { return _interopWorksheet.Name; }
            set { _interopWorksheet.Name = value; }
        }

        public ListObjectManager ListObjects
        {
            get
            {
                return this.ListObjectManager;
            }
        }

        private ListObjectManager ListObjectManager
        {
            get
            {
                if (_listObjectManager == null)
                { _listObjectManager = new ListObjectManager(_excelManager, this); }
                return _listObjectManager;
            }
        }


    }
}
