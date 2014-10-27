using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using WannaApp.Excel.Attributes;

namespace WannaApp.Excel.Helpers.MappingHelpers
{
    internal class MappingInfoManager<T>
    {

        private List<MappingInfo> _mappingInfo;

        internal List<MappingInfo> MappingInfo { get { return _mappingInfo; } }

        internal List<MappingInfo> GetMappingInfo()
        {
            ResetManager();
            BuildMappingInfo();
            return _mappingInfo.OrderBy(mp => mp.ColumnIndex).ToList();
        }

        private void BuildMappingInfo()
        {
            GetMappingForProperties();
            SetColumnIndexes();
        }

        private void SetColumnIndexes()
        {
            int index = 0;
            foreach (var mappingInfo in _mappingInfo.OrderBy(m => m.ColumnIndex))
            {
                mappingInfo.ColumnIndex = index;
                if (mappingInfo.IsDynamicRange) { index = index + mappingInfo.DynamicColumnNames.Count; }
                else { index++; }
            }
            
        }

        private void ResetManager()
        {
            _mappingInfo = new List<MappingInfo>();
        }

        private void GetMappingForProperties()
        {
            foreach (PropertyInfo property in typeof(T).GetProperties((BindingFlags.Public | BindingFlags.Instance | BindingFlags.GetProperty)))
            {
                if (GetIgnore(property) == false)
                {
                    _mappingInfo.Add(new MappingInfo
                    {
                        ColumnIndex = GetOrderIndex(property),
                        ColumnName = GetColumnName(property),
                        PropertyName = property.Name,
                        IsKey = GetIsKey(property),
                        Type = property.PropertyType

                    });
                }
            }
        }

        private int GetOrderIndex(PropertyInfo property)
        {
            ExcelMappingNameAttribute attribute = property.GetCustomAttributes(typeof(ExcelMappingNameAttribute), true).FirstOrDefault() as ExcelMappingNameAttribute;
            return (attribute == null) ? int.MaxValue : attribute.OrderIndex;
        }

        private bool GetIsKey(PropertyInfo property)
        {
            ExcelMappingKeyAttribute isKey = property.GetCustomAttributes(typeof(ExcelMappingKeyAttribute), true).FirstOrDefault() as ExcelMappingKeyAttribute;
            return (isKey != null);
        }

        private bool GetIgnore(PropertyInfo property)
        {
            ExcelMappingIgnoreAttribute skip = property.GetCustomAttributes(typeof(ExcelMappingIgnoreAttribute), true).FirstOrDefault() as ExcelMappingIgnoreAttribute;
            return (skip != null);
        }

        private string GetColumnName(PropertyInfo property)
        {
            ExcelMappingNameAttribute attribute = property.GetCustomAttributes(typeof(ExcelMappingNameAttribute), true).FirstOrDefault() as ExcelMappingNameAttribute;
            return (attribute == null || attribute.Name == null) ? property.Name : attribute.Name;
        }

        internal void SetDynamicColumnsFor(string propertyName, List<string> columnNames)
        {
            MappingInfo mappingInfo = GetMappingForProperty(propertyName);
            mappingInfo.DynamicColumnNames = columnNames;
            SetColumnIndexes();
        }

        private MappingInfo GetMappingForProperty(string propertyName)
        {
            return (from mp in _mappingInfo where mp.PropertyName == propertyName select mp).FirstOrDefault();
        }
    }
}
