using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WannaApp.Excel.ExcelObjects;
using WannaApp.Excel.Helpers;

namespace WannaApp.ExcelManagers.SettingContainers
{
    public class ValidationInfo
    {
        private string _validationFormula;
        private string _errorTitle;
        private string _errorText;
        public ValidationInfo()
        {

        }

        public string ValidationFormula()
        {
            return _validationFormula;
        }

        public ValidationInfo ValidationFormula(List<String> values)
        {
            ConvertValuesToFormula(values);
            return this;
        }

        private void ConvertValuesToFormula(List<string> values)
        {
            var stringbuilder = new StringBuilder();

            values.ForEach(v => stringbuilder.AppendFormat("{0}{1}",
                                        v,
                                        ApplicationSettings.ValidationFormulaSeparator
                                        )
                          );

            _validationFormula = stringbuilder.ToString().TrimEnd(ApplicationSettings.ValidationFormulaSeparator.ToCharArray());
        }

        public ValidationInfo ValidationFormula(ExcelRange range)
        {
            ConvertRangeToFormula();
            return this;
        }

        private void ConvertRangeToFormula()
        {
            _validationFormula = "=range"; //ToDo:
        }

        public string ErrorTitle()
        {
            return _errorTitle;
        }

        public ValidationInfo ErrorTitle(string errorTitle)
        {
            _errorTitle = errorTitle;
            return this;
        }

        public string ErrorText()
        {
            return _errorText;
        }

        public ValidationInfo ErrorText(string errorText)
        {
            _errorText = errorText;
            return this;
        }
    }
}
