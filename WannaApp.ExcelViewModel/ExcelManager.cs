using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using WannaApp.ExcelViewModel.ExcelElements;
using WannaApp.ExcelViewModel.Interfaces;

namespace WannaApp.ExcelViewModel
{
    public class ExcelManager : IExcelManager
    {
        public ExcelManager(_Application excel)
        {
            this._excelApplication = new ExcelApplication(excel);
        }

        private WorkbookManager _workbookManager;
        private WorkbookManager WorkbookManager
        {
            get
            {
                if (_workbookManager == null) { WorkbookManager = new WorkbookManager(this.ExcelApplication); }
                return _workbookManager;
            }
            set { _workbookManager = value; }
        }

        private WorksheetManager _worksheetManager;
        private WorksheetManager WorksheetManager
        {
            get
            {
                if (_worksheetManager == null) { _worksheetManager = new WorksheetManager(this); }
                return _worksheetManager;
            }
            set
            {
                _worksheetManager = value;
            }
        }

        private StylingManager _stylingManager;
        private StylingManager StylingManager
        {
            get
            {
                if (_stylingManager == null) { _stylingManager = new StylingManager(this); }
                return _stylingManager;
            }
            set { _stylingManager = value; }
        }

        private DataCollectionManager _dataCollectionManager;
        private DataCollectionManager DataCollectionManager
        {
            get
            {
                if (_dataCollectionManager == null) { _dataCollectionManager = new DataCollectionManager(this); }
                return _dataCollectionManager;
            }
            set
            {
                _dataCollectionManager = value;
            }
        }

        private ExcelApplication _excelApplication;
        public ExcelApplication ExcelApplication
        {
            get
            {
                return _excelApplication;
            }
        }
        
        public WorkbookManager Workbooks
        {
            get
            {
                return this.WorkbookManager;
            }
        }

        public WorksheetManager Sheets
        {
            get
            {
                return this.WorksheetManager;
            }
        }

        public ExcelWorkbook AddWorkbook()
        {
            var result = _excelApplication.InteropApplication.Workbooks.Add();
            return new ExcelWorkbook(this,result);
        }

        //public object GetListObjects()
        //{
        //    var result = new List<ListObject>();
        //    var listObjects = _excelApplication.InteropApplication.Workbooks[0]

        //}

        public ExcelWorksheet GetWorkSheetByName(string name)
        {
           return  WorksheetManager.GetWorkSheetByName(name);
        }
    }
}
