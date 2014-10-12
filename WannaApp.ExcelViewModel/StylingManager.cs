using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WannaApp.ExcelViewModel.Interfaces;
using WannaApp.ExcelViewModel.StylingElements;

namespace WannaApp.ExcelViewModel
{
    internal class StylingManager
    {
         private IExcelManager _excelManager;
         internal StylingManager(IExcelManager excelManager)
        {
            this._excelManager = excelManager;
        }

         internal Dictionary<string,TableStyle> TableStyles { get; set; }
         internal Dictionary<string, ColumnStyle> ColumnStyles { get; set; }
         internal Dictionary<string, RangeStyle> RangeStyles { get; set; }

    }
}
