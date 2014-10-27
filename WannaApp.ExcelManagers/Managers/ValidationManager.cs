using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WannaApp.Excel.ExcelObjects;
using WannaApp.ExcelManagers.SettingContainers;
using WannaApp.Excel.Extensions;

namespace WannaApp.ExcelManagers.Managers
{
    public class ValidationManager
    {
        private ExcelRange _range;
        private ValidationInfo _settings;
        public ValidationManager()
        {
            SetDefaults();
            
        }

        private void SetDefaults()
        {
           // no defaults yet;
        }

        public ValidationManager Range(ExcelRange range)
        {
            this._range = range;
            return this;
        }
        public ExcelRange Range()
        {
            return this._range;
        }

        public ValidationManager Settings(ValidationInfo settings )
        {
            this._settings = settings;
            return this;
        }
        public ValidationInfo Settings()
        {
            return this._settings;
        }

        public ValidationManager ApplyLookupSettings()
        {
            _range.Validation(ValidationFormula,ErrorTitle,ErrorText);
            return this;
        }

        private string ValidationFormula { get { return _settings.ValidationFormula(); } }
        private string ErrorTitle { get { return _settings.ErrorTitle(); } }
        private string ErrorText { get { return _settings.ErrorText(); } }

      

    }
}
