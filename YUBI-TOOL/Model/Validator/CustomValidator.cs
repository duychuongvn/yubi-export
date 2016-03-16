using System;
using System.ComponentModel.DataAnnotations;
using System.Text.RegularExpressions;
using YUBI_TOOL.Common;

namespace YUBI_TOOL.Model.Validator
{
    public  class CustomValidator
    {
        /// <summary>
        /// Check Time input
        /// </summary>
        /// <param name="value"> value to validate</param>
        /// <param name="vc">ValidationContext</param>
        /// <returns>ValidationResult</returns>
        public static ValidationResult ValidateTime(string value, ValidationContext vc)
        {
            if (String.IsNullOrEmpty(value))
            {
                return ValidationResult.Success;
            }


            Regex halfsizeNumberRex = new Regex("(^[0-9]{1,2}:[0-5]{1}[0-9]{0,1}$)|((^[0-9]{1,2}[0-5]{0,1}[0-9]{0,1}))");
            Match matcher = halfsizeNumberRex.Match(value);
            if (matcher.Success)
            {
                decimal time = decimal.Parse(value.Replace(":", ""));
                if (time <= 4800)
                {
                    return ValidationResult.Success;
                }
            }
            string message = null;
            var messageModel = ResourcesManager.GetMessage(Common.MessageConstant.A0027);
            if (messageModel != null)
            {
                message = messageModel.Message;
            }
            return new ValidationResult(message);
        }
        /// <summary>
        /// Check Time input
        /// </summary>
        /// <param name="value"> value to validate</param>
        /// <param name="vc">ValidationContext</param>
        /// <returns>ValidationResult</returns>
        public static ValidationResult ValidateTimeTotal(string value, ValidationContext vc)
        {
            if (String.IsNullOrEmpty(value))
            {
                return ValidationResult.Success;
            }


            Regex halfsizeNumberRex = new Regex("(^[0-9]{1,2}:[0-5]{1}[0-9]{0,1}$)|((^[0-9]{1,2}[0-5]{1}[0-9]{0,1}))");
            Match matcher = halfsizeNumberRex.Match(value);
            if (matcher.Success)
            {
                return ValidationResult.Success;
            }

            return new ValidationResult(null);
        }
    }
}
