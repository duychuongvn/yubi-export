using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.ComponentModel.DataAnnotations;
using System.Reflection;
using System.Text.RegularExpressions;
using YUBI_TOOL.Common;

namespace YUBI_TOOL.Model.Validator
{
    public class MinMaxValidator : ValidationAttribute
    {
        private readonly string NUMBER_REX = "^([-])?[0-9]+(.([0-9])+)?$";
        private string dependPropertyName;
        private string dependResourceName;
        private string formId;


        private bool isCompareLarger;

        protected override ValidationResult IsValid(object value, ValidationContext validationContext)
        {
            object model = validationContext.ObjectInstance;
            PropertyInfo dependProperty = model.GetType().GetProperty(DependPropertyName);
            if (dependProperty != null && dependProperty.CanRead)
            {
                object dependPropertyValue = dependProperty.GetValue(model, null);
                bool isValid = true;
                if (dependPropertyValue != null)
                {
                    try
                    {
                        string dependPropertyValueToString;
                        if (dependPropertyValue is Decimal)
                        {
                            dependPropertyValueToString = dependPropertyValue.ToString();
                        }
                        else
                        {
                            dependPropertyValueToString = dependPropertyValue as string;
                        }
                        if (Regex.IsMatch(dependPropertyValueToString, NUMBER_REX))
                        {
                            decimal dependPropertyValueAsDecimal = Convert.ToDecimal(dependPropertyValue);
                            //compare Larger
                            if (value != null && IsCompareLarger)
                            {
                                if (decimal.Compare(Convert.ToDecimal(value), dependPropertyValueAsDecimal) > 0)
                                {
                                    isValid = false;
                                }
                            }
                            //compare smaller
                            if (value != null && IsCompareLarger == false)
                            {
                                if (decimal.Compare(Convert.ToDecimal(value), dependPropertyValueAsDecimal) < 0)
                                {
                                    isValid = false;
                                }
                            }
                        }

                    }
                    catch
                    {
                        isValid = true;
                    }
                }

                if (!isValid)
                {
                    return new ValidationResult(null);
                }
            }
            return ValidationResult.Success;
        }

        public override string FormatErrorMessage(string name)
        {
            MessageModel messageModel = ResourcesManager.GetMessage(ErrorMessage);
            if (messageModel != null)
            {
                string thisDispName = name;
                string dependDispName = DependPropertyName;
                LanguageModel form = ResourcesManager.GetLanguageForForm(FormId);
                var thisName = ResourcesManager.GetLanguageForControlInForm(form, name);
                var dependName = ResourcesManager.GetLanguageForControlInForm(form, DependResourceName);
                if (thisName != null)
                {
                    thisDispName = thisName.Text;
                }
                if (dependName != null)
                {
                    dependDispName = dependName.Text;
                }
                if (isCompareLarger)
                {
                    return string.Format(messageModel.Message, dependDispName, thisDispName);
                }
                return string.Format(messageModel.Message, thisDispName, dependDispName);
            }
            return ErrorMessage;
        }
        public string DependPropertyName
        {
            get { return dependPropertyName; }
            set { dependPropertyName = value; }
        }

        public bool IsCompareLarger
        {
            get { return isCompareLarger; }
            set { isCompareLarger = value; }
        }
        public string DependResourceName
        {
            get { return dependResourceName; }
            set { dependResourceName = value; }
        }
        public string FormId
        {
            get { return formId; }
            set { formId = value; }
        }

    }
}
