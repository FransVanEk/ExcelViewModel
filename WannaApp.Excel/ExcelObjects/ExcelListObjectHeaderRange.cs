using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WannaApp.Excel.DataObjects;
using WannaApp.Excel.ExcelObjects;
using WannaApp.Excel.Extensions;

namespace WannaApp.Excel.ExcelObjects
{
    public class ExcelListObjectHeaderRange : ExcelRange
    {
       
        public ExcelListObjectHeaderRange(Range range): base(range)
        {
           
        }

       

    }
}
