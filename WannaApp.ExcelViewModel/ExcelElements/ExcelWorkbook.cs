using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WannaApp.ExcelViewModel.Interfaces;

namespace WannaApp.ExcelViewModel.ExcelElements
{
    public class ExcelWorkbook
    {
        private Workbook _interopWorkbook;
        private IExcelManager _excelManager;

        public ExcelWorkbook(IExcelManager excelManager, Workbook workbook)
        {
            this._excelManager = excelManager;
            this._interopWorkbook = workbook;
        }

        private WorksheetManager _worksheetManager;

        public WorksheetManager Sheets
        {
            get
            {
                if (_worksheetManager == null) { _worksheetManager = new WorksheetManager(_excelManager, this); }
                return _worksheetManager;
            }
        }

        public Workbook InteropWorkbook { get { return _interopWorkbook; } }

        public string Name
        {
            get { return _interopWorkbook.Name; }
           
        }



    }
}
