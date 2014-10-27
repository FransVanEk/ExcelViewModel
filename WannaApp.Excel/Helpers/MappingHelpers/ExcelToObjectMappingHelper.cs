using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WannaApp.Excel.Helpers.MappingHelpers;

namespace WannaApp.Excel.Helpers.MappingHelpers
{
    public class ExcelToObjectMappingHelper
    {


        private string[] headerValues;
        private int currentMappingIndex;
        private int currentHeaderIndex;
        private List<MappingInfo>  mappingInfoToExcel;
        private List<HeadersMappingInfo> mappingInfoFromExcel;

        private string CurrentHeader { get { return headerValues[currentHeaderIndex]; } }
        private MappingInfo CurrentMapping { get { return mappingInfoToExcel[currentMappingIndex]; } }
        private bool isCurrentMappingDynamicRange { get { return CurrentMapping.IsDynamicRange; } }
        private bool isCurrentMappingIndexWithinRange { get { return currentMappingIndex >= 0 && currentMappingIndex < mappingInfoToExcel.Count(); } }
        private bool isCurrentHeaderIndexWithinRange { get { return currentHeaderIndex >= 0 && currentHeaderIndex < headerValues.Length; } }

        public List<HeadersMappingInfo> Process(string[] headers, List<MappingInfo> mappingInfo)
        {
            this.headerValues= headers;
            this.mappingInfoToExcel = mappingInfo;
            mappingInfoFromExcel = new List<HeadersMappingInfo>();
            FindStartMappingIndex();
            ProcessCycle();
            return mappingInfoFromExcel;
        }

        private void ProcessCycle()
        {
            while(isCurrentMappingIndexWithinRange && isCurrentHeaderIndexWithinRange )
            {
                if (CurrentHeader == HeaderForNextMappingIndex) { currentMappingIndex++; }

                if (HeaderForCurrentMappingIndex == CurrentHeader)
                {
                    AddNewMappingInfoFromExcel();
                    currentHeaderIndex++;
                    ProcessToNextMappingMatch();
                }
                else
                {
                    currentHeaderIndex++;
                }
            }
        }

        private void ProcessToNextMappingMatch()
        {
            while(isCurrentHeaderIndexWithinRange && isCurrentMappingIndexWithinRange && CurrentHeader != HeaderForNextMappingIndex  )
            {
                if (isCurrentMappingDynamicRange && HeaderForCurrentMappingIndex != CurrentHeader)
                {
                    AddNewMappingInfoFromExcel();
                }
                currentHeaderIndex++;
            }
        }

        private void AddNewMappingInfoFromExcel()
        {
            mappingInfoFromExcel.Add(new HeadersMappingInfo { 
                            HeaderText = CurrentHeader, 
                            IsDynamicRangeProperty = CurrentMapping.IsDynamicRange, 
                            PropertyName = CurrentMapping.PropertyName,
                            ColumnIndex = currentHeaderIndex
                            });
        }

        private void FindStartMappingIndex()
        {
            while (isCurrentHeaderIndexWithinRange && HeaderForCurrentMappingIndex != CurrentHeader)
            {
                currentHeaderIndex++;
            }
        }

        private string HeaderForCurrentMappingIndex
        {
            get
            {
                return GetHeaderFor(CurrentMapping);
            }
        }

        private string GetHeaderFor(MappingInfo mappingInfo)
        {
            if (mappingInfo.IsDynamicRange) { return mappingInfo.DynamicColumnNames[0]; }
            else return mappingInfo.ColumnName;
        }

        private string HeaderForNextMappingIndex { get { return GetHeaderForNextMappingIndex(); } }

        private string GetHeaderForNextMappingIndex()
        {
            if (currentMappingIndex +1 < mappingInfoToExcel.Count) { return GetHeaderFor(mappingInfoToExcel[currentMappingIndex + 1]); }
            else { return string.Empty; }
        }
    }
}
