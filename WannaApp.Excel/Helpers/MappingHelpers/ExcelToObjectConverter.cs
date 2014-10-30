using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;

namespace WannaApp.Excel.Helpers.MappingHelpers
{
    internal class ExcelToObjectConverter<T>
    {
        private List<HeadersMappingInfo> currentMappings;
        private object[,] currentdata;
        private List<T> convertedObjects;
        private int currentRowIndex;
        private T CurrentInstance;

        private Type processingType { get { return typeof(T); } }

        public ExcelToObjectConverter<T> Process(object[,] data, List<HeadersMappingInfo> mappings)
        {
            this.currentdata = data;
            this.currentMappings = mappings;
            ProcessItems();
            return this;
        }

        private ExcelToObjectConverter<T> ProcessItems()
        {
            convertedObjects = new List<T>();
            for (int i = 1; i < currentdata.GetLength(0) + 1; i++)
            {
                currentRowIndex = i;
                GetItemForCurrentRow();
            }
            return this;
        }

        private void GetItemForCurrentRow()
        {
            T newItem = getNewInstance();
            CurrentInstance = newItem;
            FillValuesForCurrentRow();
            convertedObjects.Add(newItem);

        }

        private void FillValuesForCurrentRow()
        {
            foreach (var mapping in currentMappings)
            {
                if (mapping.IsDynamicRangeProperty)
                {
                    SetValueDynamicRange(mapping);
                }
                else
                {
                    SetValueRegular(mapping);
                }
            }
        }

        private void SetValueDynamicRange(HeadersMappingInfo mapping)
        {
            var value = GetValueFromDataArray(mapping);
            PropertyInfo prop = GetPropertyInfo(mapping);
            var propertyValue = prop.GetValue(CurrentInstance, null);
            if (value != null)
            {
                prop.PropertyType.GetMethod("Add").Invoke(propertyValue, new[] { GetConvertedValue(value, prop) });
            }
        }

        private static object GetConvertedValue(object value, PropertyInfo prop)
        {
            return Convert.ChangeType(value, prop.PropertyType.GetGenericArguments()[0]);
        }

        private void SetValueRegular(HeadersMappingInfo mapping)
        {
            var value = GetValueFromDataArray(mapping);
            PropertyInfo prop = GetPropertyInfo(mapping);
            if (prop.PropertyType == typeof(DateTime))
            {
                prop.SetValue(CurrentInstance, DateTime.FromOADate((double)Convert.ChangeType(value, typeof(double))), null);
            }
            else if (prop.PropertyType == typeof(Guid) || prop.PropertyType == typeof(Guid?))
            {
                string strvalue = (string)value;
                if (string.IsNullOrWhiteSpace(strvalue))
                {
                    prop.SetValue(CurrentInstance, null);
                }
                else
                {
                    prop.SetValue(CurrentInstance, Guid.Parse(strvalue));
                }
            }
            else
            {

                //prop.SetValue(CurrentInstance, Convert.ChangeType(value, prop.PropertyType), null);
                SetValue(prop, value);
            }
        }

        private void SetValue(PropertyInfo prop, object value)
        {
            var t = prop.PropertyType;

            if (t.IsGenericType && t.GetGenericTypeDefinition().Equals(typeof(Nullable<>)))
            {
                if (value == null)
                {
                    prop.SetValue(CurrentInstance, default(T));
                    return;
                }

                t = Nullable.GetUnderlyingType(t);
            }

            prop.SetValue(CurrentInstance, Convert.ChangeType(value, t));
        }

        public static T ChangeType<T>(object value)
        {
            var t = typeof(T);

            if (t.IsGenericType && t.GetGenericTypeDefinition().Equals(typeof(Nullable<>)))
            {
                if (value == null)
                {
                    return default(T);
                }

                t = Nullable.GetUnderlyingType(t);
            }

            return (T)Convert.ChangeType(value, t);
        }

        private PropertyInfo GetPropertyInfo(HeadersMappingInfo mapping)
        {
            return processingType.GetProperty(mapping.PropertyName);
        }

        private object GetValueFromDataArray(HeadersMappingInfo mapping)
        {
            return currentdata[currentRowIndex, mapping.ColumnIndex + 1];
        }

        private T getNewInstance()
        {
            T result = (T)Activator.CreateInstance(typeof(T));
            InstantiateDynamicRangeCollections(result);
            return result;
        }

        private void InstantiateDynamicRangeCollections(T result)
        {
            foreach (var item in GetDynamicRanges())
            {
                InstantiateDynamicRange(result, processingType.GetProperty(item.PropertyName));
            }
        }

        private void InstantiateDynamicRange(T result, System.Reflection.PropertyInfo propertyInfo)
        {
            var listType = typeof(List<>);
            var genericArgs = propertyInfo.PropertyType.GetGenericArguments();
            var concreteType = listType.MakeGenericType(genericArgs);
            var newList = Activator.CreateInstance(concreteType);
            propertyInfo.SetValue(result, newList);
        }

        private IEnumerable<HeadersMappingInfo> GetDynamicRanges()
        {
            return from item in currentMappings where item.IsDynamicRangeProperty == true select item;
        }

        public List<T> ConvertedObjects { get { return convertedObjects; } }
    }
}
