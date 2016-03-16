using System.ComponentModel.DataAnnotations;
using YUBI_TOOL.Common;
using YUBI_TOOL.Model.Validator;

namespace YUBI_TOOL.Model
{
    public class WorkDataModel : ModelBase
    {

        private decimal company_no;
        private string employee_no;
        private decimal work_date;
        private string work_date_dsp;
        private decimal? time_table_no;
        private decimal? work_type_no;
        private decimal? work_day_type_no;
        private string start_time;
        private string end_time;
        private string update_start_time;
        private string update_end_time;
        private string rest_time;
        private string contract_time;
        private string working_time;
        private string over_time;
        private string late_night_time;
        private string holiday_time;
        private string holiday_late_night_time;
        private string being_late_time;
        private string leaving_early_time;
        private decimal? work_days;
        private decimal? holiday_days;
        private decimal? paid_vacation_days;
        private string paid_vacation_time;
        private decimal? compensatory_day_off;
        private decimal? special_holidays;
        private decimal? absence_days;
        private decimal? being_late_days;
        private decimal? leaving_early_days;
        private decimal? diligence_indolence_point;
        private string memo;
        private string employeeName;
        private string post_name;
        private bool isSelected;
        private decimal post_no;
        private bool isHoliday;
        private string timeTableName;
        private string timeTable;
        private string work_type_name;
        private string employee_remarks;
        private decimal work_from;
        private decimal work_to;
        private bool isOutOfExpiration;
        private bool isHasNoOnOffDuty;
        private bool isLate;
        private bool isLeaveEarly;
        private bool isNoOnDuty;
        private bool isNoOffDuty;
        private string position;

        private WorkingDayType workingDayType;
        private WorkingType workingType;
   
        public static ValidationResult ValidateWorkType(object value, ValidationContext vc)
        {
            WorkDataModel workData = vc.ObjectInstance as WorkDataModel;
            if (workData.IsHoliday)
            {
                if (workData.work_type_no == Common.DBConstant.WORK_TYPE_ANNUAL_LEAVE
                    || workData.work_type_no == Common.DBConstant.WORK_TYPE_AW_PERMISSION
                    || workData.work_type_no == Common.DBConstant.WORK_TYPE_HALF_PERMISSION
                    || workData.work_type_no == Common.DBConstant.WORK_TYPE_MATERNITY_LEAVE
                    || workData.work_type_no == Common.DBConstant.WORK_TYPE_PERMISSION
                    || workData.work_type_no == Common.DBConstant.WORK_TYPE_SPECIAL_LEAVE)
                {
                    string message = null;
                    var messageModel = ResourcesManager.GetMessage(Common.MessageConstant.A0055);
                    if (messageModel != null)
                    {
                        message = messageModel.Message;
                        var form = ResourcesManager.GetLanguageForForm("WorkList");
                        var lblWorkType = ResourcesManager.GetLanguageForControlInForm(form, "LblWork_type");
                        if (lblWorkType != null)
                        {
                            message = string.Format(message, lblWorkType.Text);
                        }
                        else
                        {
                            message = string.Format(message, vc.MemberName);
                        }
                    }
                    return new ValidationResult(message);
                }
            }
            else
            {
                if (workData.work_type_no == Common.DBConstant.WORK_TYPE_HOLIDAY_DUTY)
                {
                    string message = null;
                    var messageModel = ResourcesManager.GetMessage(Common.MessageConstant.A0055);
                    if (messageModel != null)
                    {
                        message = messageModel.Message;
                        var form = ResourcesManager.GetLanguageForForm("WorkList");
                        var lblWorkType = ResourcesManager.GetLanguageForControlInForm(form, "LblWork_type");
                        if (lblWorkType != null)
                        {
                            message = string.Format(message, lblWorkType.Text);
                        }
                        else
                        {
                            message = string.Format(message, vc.MemberName);
                        }
                    }
                    return new ValidationResult(message);
                }
            }

            return ValidationResult.Success;
        }
        public decimal Company_no
        {
            get
            {
                return company_no;
            }
            set
            {
                if (company_no != value)
                {
                    company_no = value;
                    NotifyOfPropertyChange(() => Company_no);
                }
            }
        }

        public string Employee_no
        {
            get
            {
                return employee_no;
            }
            set
            {
                if (employee_no != value)
                {
                    employee_no = value;
                    NotifyOfPropertyChange(() => Employee_no);
                }
            }
        }

        public decimal Work_date
        {
            get
            {
                return work_date;
            }
            set
            {
                if (work_date != value)
                {
                    work_date = value;
                    NotifyOfPropertyChange(() => Work_date);
                }
            }
        }

        public decimal? Time_table_no
        {
            get
            {
                return time_table_no;
            }
            set
            {
                if (time_table_no != value)
                {
                    time_table_no = value;
                    NotifyOfPropertyChange(() => Time_table_no);
                }
            }
        }

        [CustomValidation(typeof(WorkDataModel), "ValidateWorkType", ErrorMessage = "A0055")]
        public decimal? Work_type_no
        {
            get
            {
                return work_type_no;
            }
            set
            {
                if (work_type_no != value)
                {
                    work_type_no = value;
                    NotifyOfPropertyChange(() => Work_type_no);
                }
            }
        }

        public decimal? Work_day_type_no
        {
            get
            {
                return work_day_type_no;
            }
            set
            {
                if (work_day_type_no != value)
                {
                    work_day_type_no = value;
                    NotifyOfPropertyChange(() => Work_day_type_no);
                }
            }
        }

        public string Start_time
        {
            get
            {
                return start_time;
            }
            set
            {
                if (start_time != value)
                {
                    start_time = value;
                    NotifyOfPropertyChange(() => Start_time);
                }
            }
        }

        public string End_time
        {
            get
            {
                return end_time;
            }
            set
            {
                if (end_time != value)
                {
                    end_time = value;
                    NotifyOfPropertyChange(() => End_time);
                }
            }
        }
        [Display(Name = "LblUpdate_start_time")]
        [MinMaxValidator(DependPropertyName = "Update_end_time", DependResourceName = "LblUpdate_end_time", FormId = "WorkList", IsCompareLarger = true, ErrorMessage = "A0026")]
        [CustomValidation(typeof(CustomValidator), "ValidateTime", ErrorMessage = "A0027")]
        public string Update_start_time
        {
            get
            {
                return update_start_time;
            }
            set
            {
                if (update_start_time != value)
                {
                    update_start_time = value;
                    NotifyOfPropertyChange(() => Update_start_time);
                }
            }
        }

        [Display(Name = "LblUpdate_end_time", GroupName = "WorkList")]
        [MinMaxValidator(DependPropertyName = "Update_start_time", DependResourceName = "LblUpdate_start_time", FormId = "WorkList", IsCompareLarger = false, ErrorMessage = "A0026")]
        [CustomValidation(typeof(CustomValidator), "ValidateTime", ErrorMessage = "A0027")]
        public string Update_end_time
        {
            get
            {
                return update_end_time;
            }
            set
            {
                if (update_end_time != value)
                {
                    update_end_time = value;
                    NotifyOfPropertyChange(() => Update_end_time);
                }
            }
        }

        public string Rest_time
        {
            get
            {
                return rest_time;
            }
            set
            {
                if (rest_time != value)
                {
                    rest_time = value;
                    NotifyOfPropertyChange(() => Rest_time);
                }
            }
        }

        public string Contract_time
        {
            get
            {
                return contract_time;
            }
            set
            {
                if (contract_time != value)
                {
                    contract_time = value;
                    NotifyOfPropertyChange(() => Contract_time);
                }
            }
        }

        public string Working_time
        {
            get
            {
                return working_time;
            }
            set
            {
                if (working_time != value)
                {
                    working_time = value;
                    NotifyOfPropertyChange(() => Working_time);
                }
            }
        }

        public string Over_time
        {
            get
            {
                return over_time;
            }
            set
            {
                if (over_time != value)
                {
                    over_time = value;
                    NotifyOfPropertyChange(() => Over_time);
                }
            }
        }

        public string Late_night_time
        {
            get
            {
                return late_night_time;
            }
            set
            {
                if (late_night_time != value)
                {
                    late_night_time = value;
                    NotifyOfPropertyChange(() => Late_night_time);
                }
            }
        }

        public string Holiday_time
        {
            get
            {
                return holiday_time;
            }
            set
            {
                if (holiday_time != value)
                {
                    holiday_time = value;
                    NotifyOfPropertyChange(() => Holiday_time);
                }
            }
        }

        public string Holiday_late_night_time
        {
            get
            {
                return holiday_late_night_time;
            }
            set
            {
                if (holiday_late_night_time != value)
                {
                    holiday_late_night_time = value;
                    NotifyOfPropertyChange(() => Holiday_late_night_time);
                }
            }
        }

        public string Being_late_time
        {
            get
            {
                return being_late_time;
            }
            set
            {
                if (being_late_time != value)
                {
                    being_late_time = value;
                    NotifyOfPropertyChange(() => Being_late_time);
                }
            }
        }

        public string Leaving_early_time
        {
            get
            {
                return leaving_early_time;
            }
            set
            {
                if (leaving_early_time != value)
                {
                    leaving_early_time = value;
                    NotifyOfPropertyChange(() => Leaving_early_time);
                }
            }
        }

        public decimal? Work_days
        {
            get
            {
                return work_days;
            }
            set
            {
                if (work_days != value)
                {
                    work_days = value;
                    NotifyOfPropertyChange(() => Work_days);
                }
            }
        }

        public decimal? Holiday_days
        {
            get
            {
                return holiday_days;
            }
            set
            {
                if (holiday_days != value)
                {
                    holiday_days = value;
                    NotifyOfPropertyChange(() => Holiday_days);
                }
            }
        }

        public decimal? Paid_vacation_days
        {
            get
            {
                return paid_vacation_days;
            }
            set
            {
                if (paid_vacation_days != value)
                {
                    paid_vacation_days = value;
                    NotifyOfPropertyChange(() => Paid_vacation_days);
                }
            }
        }

        public string Paid_vacation_time
        {
            get
            {
                return paid_vacation_time;
            }
            set
            {
                if (paid_vacation_time != value)
                {
                    paid_vacation_time = value;
                    NotifyOfPropertyChange(() => Paid_vacation_time);
                }
            }
        }

        public decimal? Compensatory_day_off
        {
            get
            {
                return compensatory_day_off;
            }
            set
            {
                if (compensatory_day_off != value)
                {
                    compensatory_day_off = value;
                    NotifyOfPropertyChange(() => Compensatory_day_off);
                }
            }
        }

        public decimal? Special_holidays
        {
            get
            {
                return special_holidays;
            }
            set
            {
                if (special_holidays != value)
                {
                    special_holidays = value;
                    NotifyOfPropertyChange(() => Special_holidays);
                }
            }
        }

        public decimal? Absence_days
        {
            get
            {
                return absence_days;
            }
            set
            {
                if (absence_days != value)
                {
                    absence_days = value;
                    NotifyOfPropertyChange(() => Absence_days);
                }
            }
        }

        public decimal? Being_late_days
        {
            get
            {
                return being_late_days;
            }
            set
            {
                if (being_late_days != value)
                {
                    being_late_days = value;
                    NotifyOfPropertyChange(() => Being_late_days);
                }
            }
        }

        public decimal? Leaving_early_days
        {
            get
            {
                return leaving_early_days;
            }
            set
            {
                if (leaving_early_days != value)
                {
                    leaving_early_days = value;
                    NotifyOfPropertyChange(() => Leaving_early_days);
                }
            }
        }

        public decimal? Diligence_indolence_point
        {
            get
            {
                return diligence_indolence_point;
            }
            set
            {
                if (diligence_indolence_point != value)
                {
                    diligence_indolence_point = value;
                    NotifyOfPropertyChange(() => Diligence_indolence_point);
                }
            }
        }

        public string Memo
        {
            get
            {
                return memo;
            }
            set
            {
                if (memo != value)
                {
                    memo = value;
                    NotifyOfPropertyChange(() => Memo);
                }
            }
        }
        public string EmployeeName
        {
            get
            {
                return employeeName;
            }
            set
            {
                if (employeeName != value)
                {
                    employeeName = value;
                    NotifyOfPropertyChange(() => EmployeeName);
                }
            }
        }

        public string Post_name
        {
            get
            {
                if (post_name == null)
                {
                    post_name = "";
                }
                return post_name;
            }
            set
            {
                if (post_name != value)
                {
                    post_name = value;
                    NotifyOfPropertyChange(() => Post_name);
                }
            }
        }
        public bool IsSelected
        {
            get
            {
                return isSelected;
            }
            set
            {
                if (isSelected != value)
                {
                    isSelected = value;
                    NotifyOfPropertyChange(() => IsSelected);
                }
            }
        }


        public decimal Post_no
        {
            get
            {
                return post_no;
            }
            set
            {
                if (post_no != value)
                {
                    post_no = value;
                    NotifyOfPropertyChange(() => Post_no);
                }
            }
        }
        public string Work_date_dsp
        {
            get
            {
                return work_date_dsp;
            }
            set
            {
                if (work_date_dsp != value)
                {
                    work_date_dsp = value;
                    NotifyOfPropertyChange(() => Work_date_dsp);
                }
            }
        }

        public bool IsHoliday
        {
            get
            {
                return isHoliday;
            }
            set
            {
                if (isHoliday != value)
                {
                    isHoliday = value;
                    NotifyOfPropertyChange(() => IsHoliday);
                }
            }
        }
        public string TimeTableName
        {
            get
            {
                return timeTableName;
            }
            set
            {
                if (timeTableName != value)
                {
                    timeTableName = value;
                    NotifyOfPropertyChange(() => TimeTableName);
                }
            }
        }
        public string TimeTable
        {
            get
            {
                return timeTable;
            }
            set
            {
                if (timeTable != value)
                {
                    timeTable = value;
                    NotifyOfPropertyChange(() => TimeTable);
                }
            }
        }


        public string Work_type_name
        {
            get
            {
                return work_type_name;
            }
            set
            {
                if (work_type_name != value)
                {
                    work_type_name = value;
                    NotifyOfPropertyChange(() => Work_type_name);
                }
            }
        }

        public string Employee_remarks
        {
            get
            {
                return employee_remarks;
            }
            set
            {
                if (employee_remarks != value)
                {
                    employee_remarks = value;
                    NotifyOfPropertyChange(() => Employee_remarks);
                }
            }
        }

        public decimal Work_from
        {
            get
            {
                return work_from;
            }
            set
            {
                if (work_from != value)
                {
                    work_from = value;
                    NotifyOfPropertyChange(() => Work_from);
                }
            }
        }

        public decimal Work_to
        {
            get
            {
                return work_to;
            }
            set
            {
                if (work_to != value)
                {
                    work_to = value;
                    NotifyOfPropertyChange(() => Work_to);
                }
            }
        }
        public bool IsOutOfExpiration
        {
            get
            {
                return isOutOfExpiration;
            }
            set
            {
                if (isOutOfExpiration != value)
                {
                    isOutOfExpiration = value;
                    NotifyOfPropertyChange(() => IsOutOfExpiration);
                }
            }
        }
        public bool IsHasNoOnOffDuty
        {
            get
            {
                return isHasNoOnOffDuty;
            }
            set
            {
                if (isHasNoOnOffDuty != value)
                {
                    isHasNoOnOffDuty = value;
                }
            }
        }

        public string Position
        {
            get
            {
                return position;
            }
            set
            {
                if (position != value)
                {
                    position = value;
                }
            }
        }
        public WorkingDayType WorkingDayType
        {
            get
            {
                return workingDayType;
            }
            set
            {
                if (workingDayType != value)
                {
                    workingDayType = value;
                    NotifyOfPropertyChange(() => WorkingDayType);
                }
            }
        }
        public WorkingType WorkingType
        {
            get
            {
                return workingType;
            }
            set
            {
                if (workingType != value)
                {
                    workingType = value;
                    NotifyOfPropertyChange(() => WorkingType);
                }
            }
        }

        public bool IsLate
        {
            get
            {
                return isLate;
            }
            set
            {
                if (isLate != value)
                {
                    isLate = value;
                    NotifyOfPropertyChange(() => IsLate);
                }
            }
        }

        public bool IsNoOffDuty
        {
            get { return isNoOffDuty; }
            set { isNoOffDuty = value; }
        }
        public bool IsNoOnDuty
        {
            get { return isNoOnDuty; }
            set { isNoOnDuty = value; }
        }
        
        public bool IsLeaveEarly
        {
            get
            {
                return isLeaveEarly;
            }
            set
            {
                if (isLeaveEarly != value)
                {
                    isLeaveEarly = value;
                    NotifyOfPropertyChange(() => IsLeaveEarly);
                }
            }
        }


    }


}
