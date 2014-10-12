using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using WannaApp.ExcelViewModel.ExcelElements;

namespace WannaApp.ExcelViewModel.Interfaces
{
    public interface IExcelManager
    {
        ExcelApplication ExcelApplication { get;  }
    }
}
