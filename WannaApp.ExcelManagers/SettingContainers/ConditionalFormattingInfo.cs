using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WannaApp.ExcelManagers.SettingContainers
{
    /// <summary>
    /// Enumerate for the conditional formatting. This will determine which type of condition will be checked.
    /// </summary>
    public enum eCompareType
    {
        Between,
        Equal,
        NotEqual,
        NotBetween,
        LessEqual,
        Less,
        GreaterEqual,
        Greater,
        Formula

    }

    /// <summary>
    /// Class used by the ConditionalFormatterManager in order to set the conditionalformatting to a specific range.
    /// these class contains the settings and styling info for the conditionla formatting.
    /// </summary>
    public class ConditionalFormattingInfo
    {

        public ConditionalFormattingInfo()
        {
            SetDefaults();
        }

        private void SetDefaults()
        {
            this.StopIfTrue = false;
            this.Priority = 1;
        }

        /// <summary>
        /// Color of the font
        /// </summary>
        public Color FontColor { get; set; }
        /// <summary>
        /// color of the cell interior.
        /// </summary>
        public Color BackgroundColor { get; set; }

        /// <summary>
        /// value which will be checked with all CompareTypes. For CompareType.Formula use a string but without the starting '=' char.
        /// </summary>
        public object Value1 { get; set; }
        /// <summary>
        /// value only used when compaire.type is set to Between of NotBetween. Otherwise null;
        /// </summary>
        public object Value2 { get; set; }

        /// <summary>
        /// Indicator to stop checking other conditionalformatters when the result of the check is true.
        /// this implies that the order of the conditionalformatter is important.
        /// </summary>
        public bool StopIfTrue { get; set; }
        /// <summary>
        /// indicator which compaire methode should be used.
        /// </summary>
        public eCompareType CompareType { get; set; }
        /// <summary>
        /// setter for fontcolor and backgroundcolor
        /// </summary>
        /// <param name="fontColor"></param>
        /// <param name="backgroundColor"></param>
        /// <returns></returns>
        public ConditionalFormattingInfo SetColors(Color fontColor, Color backgroundColor)
        {
            this.FontColor = fontColor;
            this.BackgroundColor = backgroundColor;
            return this;
        }
        /// <summary>
        /// indicator how to add the conditionalformatting
        /// </summary>
        public int Priority { get; set; }



        public ConditionalFormattingInfo SetStopIfTrue(bool indicator)
        {
            StopIfTrue = indicator;
            return this;
        }

        public ConditionalFormattingInfo Between(object value1, object value2)
        {
            this.CompareType = eCompareType.Between;
            this.Value1 = value1;
            this.Value2 = value2;
            return this;
        }

        public ConditionalFormattingInfo Equal(object value1)
        {
            this.CompareType = eCompareType.Equal;
            this.Value1 = value1;
            this.Value2 = null;
            return this;
        }

        public ConditionalFormattingInfo NotEqual(object value1)
        {
            this.CompareType = eCompareType.NotEqual;
            this.Value1 = value1;
            this.Value2 = null;
            return this;
        }

        public ConditionalFormattingInfo NotBetween(object value1, object value2)
        {
            this.CompareType = eCompareType.NotBetween;
            this.Value1 = value1;
            this.Value2 = value2;
            return this;
        }

        public ConditionalFormattingInfo Less(object value1)
        {
            this.CompareType = eCompareType.Less;
            this.Value1 = value1;
            this.Value2 = null;
            return this;
        }

        public ConditionalFormattingInfo LessEqual(object value1)
        {
            this.CompareType = eCompareType.LessEqual;
            this.Value1 = value1;
            this.Value2 = null;
            return this;
        }

        public ConditionalFormattingInfo GreaterEqual(object value1)
        {
            this.CompareType = eCompareType.GreaterEqual;
            this.Value1 = value1;
            this.Value2 = null;
            return this;
        }

        public ConditionalFormattingInfo Greater(object value1)
        {
            this.CompareType = eCompareType.Greater;
            this.Value1 = value1;
            this.Value2 = null;
            return this;
        }

        public ConditionalFormattingInfo Formula(string formula)
        {
            this.CompareType = eCompareType.Formula;
            this.Value1 = formula;
            this.Value2 = null;
            return this;
        }

        public ConditionalFormattingInfo SetPriority(int indicator)
        {
            this.Priority = indicator;
            return this;
        }


    }
}
