using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using Caliburn.Micro;

namespace YUBI_TOOL.Model
{
    public class ModelBase : PropertyChangedBase, IDataErrorInfo, IValidatableNestedObject
    {
        private DateTime? create_date_time;
        private DateTime? update_date_time;
        public DateTime? Create_date_time
        {
            get
            {
                return create_date_time;
            }
            set
            {
                if (create_date_time != value)
                {
                    create_date_time = value;
                    NotifyOfPropertyChange(() => Create_date_time);
                }
            }
        }

        public DateTime? Update_date_time
        {
            get
            {
                return update_date_time;
            }
            set
            {
                if (update_date_time != value)
                {
                    update_date_time = value;
                    NotifyOfPropertyChange(() => Update_date_time);
                }
            }
        }
        private bool isShowError = true;

        protected virtual ICollection<IValidatableNestedObject> GetValidatableNestedObject()
        {
            List<IValidatableNestedObject> empty = new List<IValidatableNestedObject>();
            return empty;
        }

        public bool TryValidateNestedObject(ICollection<ValidationResult> validationResults)
        {
            ICollection<IValidatableNestedObject> nestedObjects = GetValidatableNestedObject();

            foreach (IValidatableNestedObject nestedObject in nestedObjects)
            {
                nestedObject.TryValidateNestedObject(validationResults);
            }

            ValidationContext vc = new ValidationContext(this, null, null);
            System.ComponentModel.DataAnnotations.Validator.TryValidateObject(this, vc, validationResults, true);

            return validationResults.Count == 0;
        }

        public bool IsValid
        {
            get
            {
                List<ValidationResult> validationResults = new List<ValidationResult>();
                bool isValid = TryValidateNestedObject(validationResults);
                return isValid;
            }
        }

        public bool IsShowError
        {
            get { return isShowError; }
            set
            {
                if (isShowError != value)
                {
                    isShowError = value;
                    NotifyOfPropertyChange(() => IsShowError);

                }
            }
        }

        public string Error
        {
            get { throw new NotImplementedException(); }
        }

        public string this[string columnName]
        {
            get
            {
                if (IsShowError)
                {
                    var validationResults = new List<ValidationResult>();
                    object propertyValue = GetType().GetProperty(columnName).GetValue(this, null);
                    if (System.ComponentModel.DataAnnotations.Validator.TryValidateProperty(propertyValue, new ValidationContext(this, null, null) { MemberName = columnName }, validationResults))
                    {
                        return String.Empty;
                    }
                    return validationResults.First().ErrorMessage;
                }
                return String.Empty;
            }
        }

    }
}
