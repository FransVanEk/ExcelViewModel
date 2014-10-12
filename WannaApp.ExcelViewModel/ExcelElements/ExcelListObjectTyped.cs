using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WannaApp.ExcelViewModel.ExcelElements
{
    public class ExcelListObject<T> : ExcelListObjectBase
    {
        

        public ExcelListObject(ListObject listObject):base(listObject)
        {
           
        }
  

    }
}
