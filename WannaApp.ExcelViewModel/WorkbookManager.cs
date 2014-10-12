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
    public class WorkbookManager
    {
        private ExcelApplication _excelApplication;
        
        public WorkbookManager(ExcelApplication excelApplication)
        {
            this._excelApplication = excelApplication;
        }

        public Workbooks GetAllWorkbooks()
        {
            return _excelApplication.InteropApplication.Workbooks;
        }

        public WorkbookManager Save(Workbook workbook)
        {
            workbook.Save();
            return this;
        }

        public WorkbookManager Save(Workbook workbook, string filename)
        {
            workbook.SaveAs(filename);
            return this;
        }
    }
}
