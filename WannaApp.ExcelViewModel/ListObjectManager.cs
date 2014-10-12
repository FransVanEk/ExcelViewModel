using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WannaApp.ExcelViewModel
{
    public class ListObjectManager
    {
        private Interfaces.IExcelManager _excelManager;
        private ExcelElements.ExcelWorksheet excelWorksheet;

        public ListObjectManager(Interfaces.IExcelManager _excelManager, ExcelElements.ExcelWorksheet excelWorksheet)
        {
            // TODO: Complete member initialization
            this._excelManager = _excelManager;
            this.excelWorksheet = excelWorksheet;
        }
    }
}
