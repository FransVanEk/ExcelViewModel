using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WannaApp.Excel.DataObjects;
using WannaApp.Excel.ExcelObjects;

namespace WannaApp.Excel.ExcelObjects
{
    public class ExcelListObjectDataRange : ExcelRange
    {

        public ExcelListObjectDataRange(Range range)
            : base(range)
        {

        }

    }
}
