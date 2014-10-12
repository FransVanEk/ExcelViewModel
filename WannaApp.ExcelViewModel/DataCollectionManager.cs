using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WannaApp.ExcelViewModel.Interfaces;

namespace WannaApp.ExcelViewModel
{
    internal class DataCollectionManager
    {
        private IExcelManager _excelManager;
        public DataCollectionManager(IExcelManager excelManager)
        {
            this._excelManager = excelManager;
        }
       
    }
}
