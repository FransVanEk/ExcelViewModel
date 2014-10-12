using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WannaApp.Excel.DataObjects
{
    public class ListObjectDataObject : IListObjectDataObject 

    {
        public string[] HeaderValues { get; set; }
        public object[,] DataValues { get; set; }

        public object[,] _allValues;
        public object[,] AllValues { get { FillAllValues(); return _allValues; } }

        private void FillAllValues()
        {
            _allValues = new string[DataValues.GetLength(0) + 1, DataValues.GetLength(1)];
            var numberOfColumns = DataValues.GetLength(1);
            var NumberOfDataRows = DataValues.GetLength(0);
            FillHeaders(numberOfColumns);
            Array.Copy(DataValues, 0, _allValues, numberOfColumns, numberOfColumns * NumberOfDataRows);
          
        }

        private void FillHeaders(int numberOfColumns)
        {
            for (int i = 0; i < numberOfColumns; i++)
            {
                _allValues[0, i] = HeaderValues[i]; 
            }
        }

        
    }
}
