using Microsoft.Office.Interop.Excel;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using WannaApp.Excel.DataObjects;
using WannaApp.Excel.ExcelObjects;


namespace WannaApp.Excel.Helpers.MappingHelpers
{
    public class TransferHelper<T> : IListObjectDataObject where T : class
    {

        public TransferHelper()
        {
            _mappingInfoManager = new MappingInfoManager<T>();
            SetMappingInfo();

        }

        private void SetMappingInfo()
        {
            this._mappingInfoToExcel = _mappingInfoManager.GetMappingInfo();
            CreateHeaders();
        }

        public TransferHelper<T> SetDynamicColumnsFor(string propertyName, List<string> columnNames)
        {
            _mappingInfoManager.SetDynamicColumnsFor(propertyName, columnNames);
            CreateHeaders();
            return this;
        }

        public TransferHelper<T> SetDynamicColumnsFor(PropertyInfo property, List<string> columnNames)
        {
            return SetDynamicColumnsFor(property.Name, columnNames);
        }

        private IEnumerable<T> _objectDataToExcel;
        private List<T> _objectDataFromExcel;
        private List<MappingInfo> _mappingInfoToExcel;
        private List<HeadersMappingInfo> _mappingInfoFromExcel;
        private MappingInfoManager<T> _mappingInfoManager;
        private string[] _headerValues;
        private object[,] _dataValues;

        public TransferHelper<T> TransferToExcelFormat(IEnumerable<T> data)
        {
            this._objectDataToExcel = data;
            ConvertObjectDataToExcelData();
            return this;
        }

        public List<MappingInfo> MappingInfoToExcel { get { return _mappingInfoToExcel; } }

        public TransferHelper<T> TransferFromExcelFormat(ExcelListObject retrievedList)
        {
            GetExcelDataFromListObject(retrievedList);
            _mappingInfoFromExcel = new ExcelToObjectMappingHelper().Process(_headerValues, _mappingInfoToExcel);
            CreateObjects();
            return this;
        }

        private void CreateObjects()
        {
            var excelToObjectConverter = new ExcelToObjectConverter<T>();
            _objectDataFromExcel = excelToObjectConverter.Process(_dataValues, _mappingInfoFromExcel).ConvertedObjects;

        }

        private TransferHelper<T> ConvertObjectDataToExcelData()
        {
            CreateHeaders();
            CreateDataValues();
            return this;
        }

        private void CreateDataValues()
        {
            var mapping = _mappingInfoToExcel.OrderBy(mi => mi.ColumnIndex);
            var objects = _objectDataToExcel.ToList();
            _dataValues = new object[_objectDataToExcel.Count(), HeaderValues.Count()];

            for (int objectIndex = 0; objectIndex < objects.Count(); objectIndex++)
            {
                foreach (var mappingInfo in _mappingInfoToExcel.OrderBy(mi => mi.ColumnIndex))
                {
                    if (mappingInfo.IsDynamicRange)
                    {
                        var offset = 0;
                        foreach (var value in GetDynamicValues(mappingInfo, objects[objectIndex]))
                        {
                            _dataValues[objectIndex, mappingInfo.ColumnIndex + offset] = value;
                            offset++;
                        }
                    }
                    else
                    {
                        _dataValues[objectIndex, mappingInfo.ColumnIndex] = GetValue(mappingInfo, objects[objectIndex]);
                    }
                }

            }
        }

        private IEnumerable<object> GetDynamicValues(MappingInfo mappingInfo, T dataObject)
        {
            List<object> result = new List<object>();
            IEnumerable values = dataObject.GetType().GetProperty(mappingInfo.PropertyName).GetValue(dataObject, null) as IEnumerable;

            if(values != null) { 
                   foreach (object element in values)
        {
                result.Add(element);            
        }
            }
            return result;
           
        }

        private object GetValue(MappingInfo mappingInfo, T dataObject)
        {
            if (mappingInfo.Type == typeof(Guid))
            {
                return dataObject.GetType().GetProperty(mappingInfo.PropertyName).GetValue(dataObject, null).ToString();
            }
            else
            {
                return dataObject.GetType().GetProperty(mappingInfo.PropertyName).GetValue(dataObject, null);
            }
        }

        private void CreateHeaders()
        {
            var result = new List<string>();

            foreach (var item in _mappingInfoToExcel.OrderBy(mi => mi.ColumnIndex))
            {
                result.AddRange(GetHeadersFor(item));
            }

            _headerValues = result.ToArray();
        }

        private IEnumerable<string> GetHeadersFor(MappingInfo item)
        {
            var result = new List<string>();
            if (item.IsDynamicRange == false) { result.Add(item.ColumnName); }
            else { result.AddRange(item.DynamicColumnNames); }
            return result;
        }

        public object[,] AllValues
        {
            get { return GetAllValues(); }
        }

        private object[,] GetAllValues()
        {
            var result = new object[DataValues.GetLength(0) + 1, HeaderValues.Length];
            CoyyHeadersInToResult(result);
            CopyDataToResult(result);
            return result;
        }

        private void CopyDataToResult(object[,] result)
        {
            for (int row = 0; row < DataValues.GetLength(0); row++)
            {
                for (int col = 0; col < DataValues.GetLength(1); col++)
                {
                    result[row + 1, col] = DataValues[row, col];
                }

            }
        }

        private object[,] CoyyHeadersInToResult(object[,] result)
        {

            for (int i = 0; i < HeaderValues.Count(); i++)
            {
                result[0, i] = HeaderValues[i];
            }
            return result;
        }

        public object[,] DataValues
        {
            get { return _dataValues; }
        }

        public string[] HeaderValues
        {
            get { return _headerValues; }
        }

        private void GetExcelDataFromListObject(ExcelListObject retrievedList)
        {
            this._headerValues = retrievedList.Headers;
            this._dataValues = retrievedList.dataValues;
        }

        public IEnumerable<T> GetObjectsFromExcel()
        {
            return _objectDataFromExcel;
        }
    }
}
