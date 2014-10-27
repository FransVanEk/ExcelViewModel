//using System;
//using System.Collections.Generic;
//using System.Linq;
//using WannaApp.Excel.ExcelObjects;
//using WannaApp.ExcelManagers.SettingContainers;

//namespace WannaApp.ExcelManagers.Managers
//{
//    public class ConditionalFormatManager
//    {
//        private ExcelRange _range;
//        private ConditionalFormattingInfo _currentFormattingInfo;

//        /// <summary>
//        /// List of conditional formatting settings. in case of a table these conditions 
//        /// will be applied after the styling for the table 
//        /// </summary>
//        public List<ConditionalFormattingInfo> FormattingInfo { get; set; }

//        /// <summary>
//        /// Manager for apply conditional formatting on ranges. Settings can be stored but will not be enforced directly.
//        /// Only after calling Applyformatting the formaating wil be enforced to the range.
//        /// </summary>
//        /// <param name="range">One or more cells combined in a range to which the formatting must be applied</param>
//        public ConditionalFormatManager(ExcelRange range)
//        {
//            InitializeClass();
//            SetRange(range);
//        }

//        private void InitializeClass()
//        {
//            this.FormattingInfo = new List<ConditionalFormattingInfo>();
//        }

//        /// <summary>
//        /// One or more cells combined in a range to which the formatting must be applied
//        /// </summary>
//        /// <param name="range">combined cells in a range</param>
//        /// <returns></returns>
//        public ConditionalFormatManager SetRange(ExcelRange range)
//        {
//            this._range = range;
//            return this;
//        }
//        /// <summary>
//        /// After storing the conditional formatting, the settings will be enforced by calling this method.
//        /// </summary>
//        /// <returns></returns>
//        public ConditionalFormatManager ApplyFormatting()
//        {
//            FormattingInfo
//                    .OrderBy(x => x.Priority)
//                    .ToList()
//                    .ForEach(cf => ApplyFormatting(cf));
//            return this;
//        }

//        private void ApplyFormatting(ConditionalFormattingInfo formattingInfo)
//        {
//            this._currentFormattingInfo = formattingInfo;
//            FormatCondition formatCondition = GetFormatCondition();
//            if (formatCondition != null)
//            {
//                SetStyling(formatCondition);
//                if (OrderNecessary)
//                {
//                    OrderConditionalFormatting(formatCondition);
//                }
//            }
//        }

//        private bool OrderNecessary
//        {
//            get
//            {
//                return _currentFormattingInfo.Priority < NumberOfCurrentItems && _currentFormattingInfo.Priority > 0;
//            }
//        }

//        private void OrderConditionalFormatting(FormatCondition conditinalformatting)
//        {
//            conditinalformatting.Priority = _currentFormattingInfo.Priority;


//        }

//        private int NumberOfCurrentItems
//        {
//            get
//            {
//                return _range.FormatConditions.Count;
//            }
//        }

//        private FormatCondition GetFormatCondition()
//        {
//            FormatCondition formatCondition;

//            switch (_currentFormattingInfo.CompareType)
//            {
//                case eCompareType.Between:
//                    formatCondition = GetBetweenFormattingCurrent();
//                    break;
//                case eCompareType.Equal:
//                    formatCondition = GetEqualFormattingCurrent();
//                    break;
//                case eCompareType.NotEqual:
//                    formatCondition = GetNotEqualFormattingCurrent();
//                    break;
//                case eCompareType.NotBetween:
//                    formatCondition = GetNotBetweenFormattingCurrent();
//                    break;
//                case eCompareType.LessEqual:
//                    formatCondition = GetLessEqualFormattingCurrent();
//                    break;
//                case eCompareType.Less:
//                    formatCondition = GetLessFormattingCurrent();
//                    break;
//                case eCompareType.GreaterEqual:
//                    formatCondition = GetGreaterEqualFormattingCurrent();
//                    break;
//                case eCompareType.Greater:
//                    formatCondition = GetGreaterFormattingCurrent();
//                    break;
//                case eCompareType.Formula:
//                    formatCondition = GetFormulaFormattingCurrent();
//                    break;

//                default:
//                    formatCondition = null;
//                    break;

//            }
//            return formatCondition;
//        }

//        private FormatCondition GetFormulaFormattingCurrent()
//        {
//            return _range.FormatConditions.Add(XlFormatConditionType.xlExpression, Type.Missing, GetValue1Formula(), Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

//        }

//        private FormatCondition GetNotEqualFormattingCurrent()
//        {
//            return _range.FormatConditions.Add(XlFormatConditionType.xlCellValue, XlFormatConditionOperator.xlEqual, GetValue1Formula());

//        }

//        private FormatCondition GetNotBetweenFormattingCurrent()
//        {
//            return _range.FormatConditions.Add(XlFormatConditionType.xlCellValue, XlFormatConditionOperator.xlNotBetween, GetValue1Formula(), GetValue2Formula());
//        }

//        private FormatCondition GetLessEqualFormattingCurrent()
//        {
//            return _range.FormatConditions.Add(XlFormatConditionType.xlCellValue, XlFormatConditionOperator.xlLessEqual, GetValue1Formula());
//        }

//        private FormatCondition GetLessFormattingCurrent()
//        {
//            return _range.FormatConditions.Add(XlFormatConditionType.xlCellValue, XlFormatConditionOperator.xlLess, GetValue1Formula());
//        }

//        private FormatCondition GetGreaterEqualFormattingCurrent()
//        {
//            return _range.FormatConditions.Add(XlFormatConditionType.xlCellValue, XlFormatConditionOperator.xlGreaterEqual, GetValue1Formula());
//        }

//        private FormatCondition GetGreaterFormattingCurrent()
//        {
//            return _range.FormatConditions.Add(XlFormatConditionType.xlCellValue, XlFormatConditionOperator.xlGreater, GetValue1Formula());
//        }

//        private FormatCondition GetEqualFormattingCurrent()
//        {
//            return _range.FormatConditions.Add(XlFormatConditionType.xlCellValue, XlFormatConditionOperator.xlEqual, GetValue1Formula());
//        }

//        private FormatCondition GetBetweenFormattingCurrent()
//        {
//            return _range.FormatConditions.Add(XlFormatConditionType.xlCellValue, XlFormatConditionOperator.xlBetween, GetValue1Formula(), GetValue2Formula());
//        }

//        /// <summary>
//        /// Fucntion whish enforces the styling for the range when the condition is met.
//        /// </summary>
//        /// <param name="conditinalformatting"></param>
//        private void SetStyling(FormatCondition conditinalformatting)
//        {
//            //this is the spot to add additional styling possibilities to the ranges
//            if (_currentFormattingInfo.FontColor != Color.Empty)
//            {
//                conditinalformatting.Font.Color = _currentFormattingInfo.FontColor;
//            }
//            if (_currentFormattingInfo.BackgroundColor != Color.Empty)
//            {
//                conditinalformatting.Interior.Color = _currentFormattingInfo.BackgroundColor;
//            }
//            conditinalformatting.StopIfTrue = _currentFormattingInfo.StopIfTrue;

//        }

//        private object GetValue1Formula()
//        {
//            return GetFormula(_currentFormattingInfo.Value1);
//        }

//        private object GetValue2Formula()
//        {
//            return GetFormula(_currentFormattingInfo.Value2);
//        }

//        private object GetFormula(object value)
//        {
//            if (value == null)
//            {
//                return Type.Missing;
//            }
//            else
//            {
//                return string.Format("={0}", value);
//            }
//        }

//        /// <summary>
//        /// Adds a new formatting setting to the collection of formatters. 
//        /// Will be applied to the range after calling the methiod applyFormatting.
//        /// </summary>
//        /// <param name="formattingInfo"></param>
//        /// <returns></returns>
//        public ConditionalFormatManager Add(ConditionalFormattingInfo formattingInfo)
//        {
//            this.FormattingInfo.Add(formattingInfo);
//            return this;
//        }



//    }
//}
