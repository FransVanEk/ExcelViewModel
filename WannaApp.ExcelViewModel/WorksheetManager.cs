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
    public class WorksheetManager
    {

        private IExcelManager _excelManager;
        private ExcelWorkbook _excelWorkBook;
        private WorkbookManager workbookManager;

        private Sheets Worksheets
        {
            get
            {
                if (_excelWorkBook != null) { return _excelWorkBook.InteropWorkbook.Worksheets; }
                else { return _excelManager.ExcelApplication.InteropApplication.Worksheets; }
            }
        }

        public WorksheetManager(IExcelManager excelManager)
        {
            this._excelManager = excelManager;
        }
        public WorksheetManager(IExcelManager excelManager, ExcelWorkbook workbook)
        {
            this._excelManager = excelManager;
            this._excelWorkBook = workbook;
        }

        public ExcelWorksheet AddWorksheet(string name)
        {
            Worksheet newSheet = Worksheets.Add();
            newSheet.Name = name;
            return new ExcelWorksheet(_excelManager,  newSheet);
        }

        public ExcelWorksheet AddWorksheet(string name, Worksheet worksheetBefore)
        {
            Worksheet newSheet = Worksheets.Add(worksheetBefore);
            newSheet.Name = name;
            return new ExcelWorksheet(_excelManager, newSheet);
        }

        public List<ExcelWorksheet> AllSheets()
        {
            var result = new List<ExcelWorksheet>();

            foreach (var worksheet in Worksheets)
            {
                result.Add(new ExcelWorksheet(_excelManager,  (Worksheet)worksheet));
            }
            return result;
        }

        internal ExcelWorksheet GetWorkSheetByName(string name)
        {
            var result = AllSheets().Where(ws => ws.Name == name).FirstOrDefault();
            if (result != null) {return result; }
            return result;
        }

        public void Remove(ExcelWorksheet worksheet)
        {
            if (worksheet != null)
            {
                worksheet.InteropWorkSheet.Delete();
            }
        }
    }
}
