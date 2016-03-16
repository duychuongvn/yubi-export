using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Windows;
using Caliburn.Micro;
using Microsoft.Office.Interop.Excel;
using YUBI_TOOL.Common;
using YUBI_TOOL.Model;
using YUBI_TOOL.Service;
using Excel = Microsoft.Office.Interop.Excel;

namespace YUBI_TOOL.ViewModel
{
    public class AttendanceListViewModel : ViewModelBase
    {
        private const string FORM_ID = "WorkList";

        private readonly IWorkDataService workDataService;
        private readonly IEmployeeService employeeService;
        private readonly ITimeTableService timeTableService;
        private readonly IHolidayService holidayService;
        private readonly IMonthlyService monthlyService;
        private string employeeNo;
        private decimal companyNo;
        private decimal post_no;
        private decimal currentYearMonth;
        private string dateFormat;
        private string employeeName;
        private EmployeeModel employeeModel;
        private List<WorkDataModel> workDataList;
        private List<SelectItemModel> dayTypeList;
        private List<TimeTableModel> timeTableList;
        private List<HolidayModel> holidayList;
        private WorkDataModel selectedWorkData;
        private LanguageModel lblAttendenceTitle;
        private LanguageModel lblAttendenceName;
        private LanguageModel lblAttendencePeriod;
        private LanguageModel lblPrev;
        private LanguageModel lblNext;
        private LanguageModel lblMessageArea;
        private LanguageModel lblMessageText;
        private LanguageModel lblEmployeeName;
        private LanguageModel lblPeriod;
        private LanguageModel lblPersonalExport;
        private LanguageModel lblUpdate;
        private LanguageModel lblCancel;
        private string[] headerText;
        private bool isFocused;
        private string period;
        private string workDateFormat = "MD";
        private WorkDataModel copiedWorkDataModel;
        private WorkDataModel pasteWorkDataModel;

        private string deletePropertyName;
        private bool isChanged = false;
        private bool isCopyMode = false;
        private bool hasAutoChangeContactTime = false;
        private int unit_minutes = 0;
        private Excel.Application excelApp = null;
        private Excel.Workbook workBook;
        private Excel.Workbook tempWorkBook;
        public AttendanceListViewModel()
        {
            workDataService = IoC.Get<IWorkDataService>();
            employeeService = IoC.Get<IEmployeeService>();
            holidayService = IoC.Get<IHolidayService>();
            timeTableService = IoC.Get<ITimeTableService>();
            monthlyService = IoC.Get<IMonthlyService>();
            this.PropertyChanged += new System.ComponentModel.PropertyChangedEventHandler(AttendanceListViewModel_PropertyChanged);
        }

        private void AttendanceListViewModel_PropertyChanged(object sender, System.ComponentModel.PropertyChangedEventArgs e)
        {
            if (e.PropertyName == GetPropertyName(() => PasteWorkDataModel))
            {
                DoPaste();
            }
            else if (e.PropertyName == GetPropertyName<AttendanceListViewModel, string>(x => x.DeletePropertyName)
                && !string.IsNullOrEmpty(DeletePropertyName)
                && SelectedWorkData != null
                && DeletePropertyName != GetPropertyName<WorkDataModel, decimal?>(x => x.Work_type_no)
                && DeletePropertyName != GetPropertyName<WorkDataModel, decimal?>(x => x.Time_table_no)
                && DeletePropertyName != GetPropertyName<WorkDataModel, string>(x => x.Paid_vacation_time))
            {
                if (DeletePropertyName == GetPropertyName<WorkDataModel, string>(x => x.Work_date_dsp))
                {

                    SelectedWorkData.Work_type_no = DBConstant.WORK_TYPE_NORMAL;
                    SelectedWorkData.Time_table_no = DBConstant.TIME_TABLE_NO_DEFAULT;
                    SelectedWorkData.Absence_days = null;
                    SelectedWorkData.Being_late_days = null;
                    SelectedWorkData.Being_late_time = null;
                    SelectedWorkData.Compensatory_day_off = null;
                    SelectedWorkData.Contract_time = null;
                    SelectedWorkData.Diligence_indolence_point = null;
                    SelectedWorkData.End_time = null;
                    SelectedWorkData.Holiday_days = null;
                    SelectedWorkData.Holiday_late_night_time = null;
                    SelectedWorkData.Holiday_time = null;
                    SelectedWorkData.Late_night_time = null;
                    SelectedWorkData.Leaving_early_days = null;
                    SelectedWorkData.Leaving_early_time = null;
                    SelectedWorkData.Memo = null;
                    SelectedWorkData.Over_time = null;
                    SelectedWorkData.Paid_vacation_days = null;
                    SelectedWorkData.Rest_time = null;
                    SelectedWorkData.Special_holidays = null;
                    SelectedWorkData.Start_time = null;
                    SelectedWorkData.Update_end_time = null;
                    SelectedWorkData.Update_start_time = null;
                    SelectedWorkData.Work_days = null;
                    SelectedWorkData.Working_time = null;


                }
                else
                {
                    selectedWorkData.GetType().GetProperty(DeletePropertyName).SetValue(selectedWorkData, null, null);
                }
                DeletePropertyName = null;
            }
        }
        public void Init(string employeeNo, decimal companyNo, decimal postNo, int unit_minutes, decimal yearMonth)
        {
            this.employeeNo = employeeNo;
            this.companyNo = companyNo;
            this.post_no = postNo;
            this.currentYearMonth = yearMonth;
            decimal lastDayOfMonth = CommonUtil.GetLastDayOfMonth(currentYearMonth);
            decimal firstDayOfMonth = CommonUtil.GetFirstDayOfMonth(currentYearMonth);
            this.employeeModel = employeeService.GetEmployee(companyNo, post_no, employeeNo, firstDayOfMonth, lastDayOfMonth);
            if (employeeModel != null)
            {
                EmployeeName = CommonUtil.GetFullName(employeeModel.Emsize_first_name, employeeModel.Emsize_last_name);
            }
            this.unit_minutes = unit_minutes;
        }

        private void GetData()
        {
            if (WorkDataList != null)
            {
                CommonUtil.ClearSortDirection(WorkDataList);
            }
            decimal firstDayOfMonth = CommonUtil.GetFirstDayOfMonth(currentYearMonth);
            decimal lastDayOfMonth = CommonUtil.GetLastDayOfMonth(currentYearMonth);
            Period = CommonUtil.GetDateAsString(firstDayOfMonth, dateFormat, false) + " - " + CommonUtil.GetDateAsString(CommonUtil.GetLastDayOfMonth(currentYearMonth), dateFormat, false);

            workDataService.CreateWorkDataInMonth(employeeModel, firstDayOfMonth);
            WorkDataList = workDataService.SearchWorkDataListByEmployee(companyNo, post_no, employeeNo, firstDayOfMonth, lastDayOfMonth);
            SelectedWorkData = null;
            hasAutoChangeContactTime = false;
            ParseData();
            DoValidate();
            isChanged = false;

        }

        public bool IsEmployeeInExpiration(decimal date)
        {
            bool isValid = false;
            if (employeeModel.Expiration_from <= date && employeeModel.Expiration_to >= date)
            {
                isValid = true;
            }
            return isValid;
        }

        public void ChangeWorkDayType(WorkDataModel workDataModel)
        {
            if (workDataModel != null)
            {
                if (workDataModel.Work_day_type_no <= DBConstant.WORK_DAY_TYPE_NORMAL)
                {
                    workDataModel.Work_day_type_no = DBConstant.WORK_DAY_TYPE_NORMAL_HOLIDAY;
                    workDataModel.IsHoliday = true;
                    workDataModel.WorkingDayType = WorkingDayType.NormalHoliday;
                }
                else if (workDataModel.Work_day_type_no <= DBConstant.WORK_DAY_TYPE_NORMAL_HOLIDAY)
                {
                    workDataModel.Work_day_type_no = DBConstant.WORK_DAY_TYPE_NATIONAL_HOLIDAY;
                    workDataModel.IsHoliday = true;
                    workDataModel.WorkingDayType = WorkingDayType.NationalHoliday;
                }
                else
                {
                    workDataModel.Work_day_type_no = DBConstant.WORK_DAY_TYPE_NORMAL;
                    workDataModel.IsHoliday = false;
                    workDataModel.WorkingDayType = WorkingDayType.Normal;
                }
                workDataModel.Work_type_no = DBConstant.WORK_TYPE_NORMAL;
                workDataModel.WorkingType = WorkingType.Unknown;
            }
        }
        private void ParseData()
        {
            foreach (var workData in workDataList)
            {
                workData.Work_date_dsp = CommonUtil.GetDateAsString(workData.Work_date, workDateFormat, true);
                workData.WorkingDayType = CommonUtil.GetWorkingDayType(workData, employeeModel, holidayList);

                if (workData.WorkingDayType == WorkingDayType.NationalHoliday || workData.WorkingDayType == WorkingDayType.NormalHoliday)
                {
                    workData.IsHoliday = true;
                }

                if (workData.Work_date < employeeModel.Expiration_from || workData.Work_date > employeeModel.Expiration_to)
                {
                    workData.IsOutOfExpiration = true;
                }
                else
                {
                    workData.IsOutOfExpiration = false;
                }

                decimal? contactTime = CommonUtil.ToNullableDecimal(workData.Contract_time);
                TimeTableModel timeTable = timeTableList.Find(x => x.Time_table_no == workData.Time_table_no);
                if (timeTable != null)
                {
                    workData.TimeTable = CommonUtil.GetTimeShiftDsp(timeTable);
                    if ((contactTime == null || contactTime == 0)
                        && (workData.Work_type_no == DBConstant.WORK_TYPE_NORMAL || workData.Work_type_no == DBConstant.WORK_TYPE_HALF_PERMISSION)
                        && !workData.IsHoliday)
                    {

                        decimal contact_time = CommonUtil.GetContactTime(timeTable);
                        if (workData.Work_type_no == DBConstant.WORK_TYPE_HALF_PERMISSION)
                        {
                            contact_time = contact_time / 2;
                        }

                        workData.Contract_time = contact_time.ToString();
                        hasAutoChangeContactTime = true;
                    }
                }
                workData.PropertyChanged -= new System.ComponentModel.PropertyChangedEventHandler(WorkData_PropertyChanged);
                workData.PropertyChanged += new System.ComponentModel.PropertyChangedEventHandler(WorkData_PropertyChanged);
            }
        }

        private bool DoCheckInput(WorkDataModel workModel)
        {
            bool isValid = true;
            List<string> messages = GetMessage(workModel);
            if (messages.Count > 0)
            {
                isValid = false;
                Message = new MessageModel();
                Message.Foreground = ResourcesManager.GetForeground(ResourcesManager.KEY_COLOR_MESSAGE_ERROR);
                Message.Background = ResourcesManager.GetBackground(ResourcesManager.KEY_COLOR_MESSAGE_ERROR);
                foreach (var error in messages)
                {
                    Message.Message += error;
                    break;
                }
            }
            else
            {
                Message = ResourcesManager.GetMessage(MessageConstant.I1001);
            }
            return isValid;
        }
        private void WorkData_PropertyChanged(object sender, System.ComponentModel.PropertyChangedEventArgs e)
        {
            WorkDataModel workData = sender as WorkDataModel;
            isChanged = true;
            if (isCopyMode || !DoCheckInput(workData))
            {
                return;
            }
            if (e.PropertyName == GetPropertyName<WorkDataModel, decimal?>(x => x.Time_table_no))
            {
                workData.Update_start_time = null;
                workData.Update_end_time = null;
                workData.Working_time = null;
                workData.Rest_time = null;
                workData.Over_time = null;
                workData.Being_late_time = null;
                workData.Leaving_early_time = null;
                workData.Update_start_time = workData.Start_time;
                workData.Update_end_time = workData.End_time;
                TimeTableModel timeTable = timeTableList.Find(x => x.Time_table_no == workData.Time_table_no);
                if (timeTable != null)
                {
                    workData.Work_from = timeTable.Work_from;
                    workData.Work_to = timeTable.Work_to;
                }
                else
                {
                    workData.Work_from = 0;
                    workData.Work_to = 0;
                }
                if (workData.Work_type_no == DBConstant.WORK_TYPE_NORMAL && !workData.IsHoliday)
                {
                    workData.Contract_time = CommonUtil.GetContactTime(timeTable).ToString();
                }
                else
                {
                    workData.Contract_time = "0";
                }
            }
            else if (e.PropertyName == GetPropertyName<WorkDataModel, string>(x => x.Update_start_time)
               || e.PropertyName == GetPropertyName<WorkDataModel, string>(x => x.Update_end_time)
               || e.PropertyName == GetPropertyName<WorkDataModel, decimal?>(x => x.Work_type_no))
            {
                if (workData.Work_type_no == DBConstant.WORK_TYPE_NORMAL
                    || workData.Work_type_no == DBConstant.WORK_TYPE_HOLIDAY_DUTY
                    || workData.Work_type_no == DBConstant.WORK_TYPE_HALF_PERMISSION)
                {
                    var timeTable = timeTableList.Find(x => x.Time_table_no == workData.Time_table_no);
                    decimal? update_start_time = CommonUtil.ToNullableDecimal(workData.Update_start_time);
                    decimal? update_end_time = CommonUtil.ToNullableDecimal(workData.Update_end_time);
                    if (update_start_time != null && update_end_time != null)
                    {
                        if (update_end_time <= update_start_time)
                        {
                            return;
                        }
                        string restFrom = "Rest{0}_from";
                        string restTo = "Rest{0}_to";
                        decimal restTime = 0;
                        decimal overTime = 0;
                        decimal lateness = 0;
                        decimal leaveEarly = 0;

                        if (timeTable.Work_from < update_start_time)
                        {
                            lateness = CommonUtil.SubTime(update_start_time.Value, timeTable.Work_from);
                        }
                        if (timeTable.Work_to > CommonUtil.ToNullableDecimal(workData.Update_end_time))
                        {
                            leaveEarly = CommonUtil.SubTime(timeTable.Work_to, update_end_time.Value);
                        }
                        for (int i = 1; i <= 10; i++)
                        {
                            decimal? restTimeFrom = (decimal?)CommonUtil.GetPropertyValue(timeTable, string.Format(restFrom, i));
                            decimal? restTimeTo = (decimal?)CommonUtil.GetPropertyValue(timeTable, string.Format(restTo, i));

                            if (restTimeFrom != null && restTimeTo != null)
                            {

                                decimal over_Time_at_start = 0;
                                decimal over_Time_at_end = 0;

                                // start work
                                if (restTimeFrom.Value < timeTable.Work_from)
                                {
                                    if (update_start_time < restTimeFrom.Value)
                                    {
                                        restTime += CommonUtil.SubTime(restTimeTo.Value, restTimeFrom.Value);
                                        over_Time_at_start = CommonUtil.SubTime(restTimeFrom.Value, update_start_time.Value);
                                    }
                                    else if (lateness == 0)
                                    {
                                        restTime += CommonUtil.SubTime(timeTable.Work_from, update_start_time.Value);
                                    }
                                }
                                // end work
                                else if (restTimeTo > timeTable.Work_to)
                                {
                                    if (update_end_time > restTimeTo.Value)
                                    {
                                        restTime += CommonUtil.SubTime(restTimeTo.Value, restTimeFrom.Value);
                                        over_Time_at_end = CommonUtil.SubTime(update_end_time.Value, restTimeTo.Value);
                                    }
                                    else if (leaveEarly == 0)
                                    {
                                        // not late
                                        restTime += CommonUtil.SubTime(update_end_time.Value, timeTable.Work_to);
                                    }
                                }
                                // rest inmindle working
                                else
                                {
                                    restTime += CommonUtil.SubTime(restTimeTo.Value, restTimeFrom.Value);
                                }

                                overTime += over_Time_at_start + over_Time_at_end;
                            }

                        }
                        if (leaveEarly > 0)
                        {
                            workData.Leaving_early_time = CommonUtil.ToString(leaveEarly);
                        }
                        else
                        {
                            workData.Leaving_early_time = null;
                        }
                        if (lateness > 0)
                        {
                            workData.Being_late_time = CommonUtil.ToString(lateness);
                        }
                        else
                        {
                            workData.Being_late_time = null;
                        }
                        workData.Rest_time = CommonUtil.ToString(restTime);
                        workData.Over_time = CommonUtil.ToString(overTime);
                        decimal working_time = CommonUtil.SubTime(timeTable.Work_to, timeTable.Work_from) - leaveEarly - lateness + overTime;
                        if (workData.Work_type_no == DBConstant.WORK_TYPE_HOLIDAY_DUTY)
                        {
                            workData.Working_time = null;
                            workData.Holiday_time = CommonUtil.ToString(working_time);
                            workData.Holiday_days = 1M;
                        }
                        else
                        {
                            workData.Holiday_time = null;
                            workData.Working_time = CommonUtil.ToString(working_time);
                            workData.Holiday_days = null;
                        }

                        if (workData.Work_type_no == DBConstant.WORK_TYPE_NORMAL && !workData.IsHoliday)
                        {
                            workData.Contract_time = CommonUtil.GetContactTime(timeTable).ToString();
                            workData.Work_days = 1M;
                        }
                        else if (workData.Work_type_no == DBConstant.WORK_TYPE_HALF_PERMISSION)
                        {
                            workData.Contract_time = CommonUtil.DivideTime(CommonUtil.GetContactTime(timeTable), 2).ToString();
                            workData.Late_night_time = null;
                            workData.Leaving_early_days = null;
                            workData.Leaving_early_time = null;
                            workData.Work_days = 0.5M;
                        }
                        else
                        {
                            workData.Contract_time = "0";
                            workData.Holiday_days = null;
                            workData.Work_days = null;
                        }
                    }
                    else
                    {
                        workData.Working_time = null;
                        workData.Rest_time = null;
                        workData.Over_time = null;
                        workData.Being_late_time = null;
                        workData.Leaving_early_time = null;
                        workData.Holiday_days = null;
                        workData.Work_days = null;
                    }
                }
                else
                {
                    workData.Contract_time = "0";
                    workData.Update_start_time = null;
                    workData.Update_end_time = null;
                    workData.Working_time = null;
                    workData.Rest_time = null;
                    workData.Over_time = null;
                    workData.Being_late_time = null;
                    workData.Leaving_early_time = null;
                }

            }
        }


        public void DoPaste()
        {
            if (copiedWorkDataModel != null && pasteWorkDataModel != null && !pasteWorkDataModel.IsHoliday
                && DoCheckInput(copiedWorkDataModel))
            {
                isCopyMode = true;
                pasteWorkDataModel.Absence_days = copiedWorkDataModel.Absence_days;
                pasteWorkDataModel.Being_late_days = copiedWorkDataModel.Being_late_days;
                pasteWorkDataModel.Being_late_time = copiedWorkDataModel.Being_late_time;
                pasteWorkDataModel.Compensatory_day_off = copiedWorkDataModel.Compensatory_day_off;
                pasteWorkDataModel.Contract_time = copiedWorkDataModel.Contract_time;
                pasteWorkDataModel.Diligence_indolence_point = copiedWorkDataModel.Diligence_indolence_point;
                pasteWorkDataModel.Employee_remarks = copiedWorkDataModel.Employee_remarks;
                pasteWorkDataModel.End_time = copiedWorkDataModel.End_time;
                pasteWorkDataModel.Holiday_days = copiedWorkDataModel.Holiday_days;
                pasteWorkDataModel.Holiday_late_night_time = copiedWorkDataModel.Holiday_late_night_time;
                pasteWorkDataModel.Holiday_time = copiedWorkDataModel.Holiday_time;
                pasteWorkDataModel.Late_night_time = copiedWorkDataModel.Late_night_time;
                pasteWorkDataModel.Leaving_early_days = copiedWorkDataModel.Leaving_early_days;
                pasteWorkDataModel.Leaving_early_time = copiedWorkDataModel.Leaving_early_time;
                pasteWorkDataModel.Memo = copiedWorkDataModel.Memo;
                pasteWorkDataModel.Over_time = copiedWorkDataModel.Over_time;
                pasteWorkDataModel.Paid_vacation_days = copiedWorkDataModel.Paid_vacation_days;
                pasteWorkDataModel.Paid_vacation_time = copiedWorkDataModel.Paid_vacation_time;
                pasteWorkDataModel.Post_name = copiedWorkDataModel.Post_name;
                pasteWorkDataModel.Post_no = copiedWorkDataModel.Post_no;
                pasteWorkDataModel.Rest_time = copiedWorkDataModel.Rest_time;
                pasteWorkDataModel.Special_holidays = copiedWorkDataModel.Special_holidays;
                pasteWorkDataModel.Start_time = copiedWorkDataModel.Start_time;
                pasteWorkDataModel.Time_table_no = copiedWorkDataModel.Time_table_no;
                pasteWorkDataModel.TimeTableName = copiedWorkDataModel.TimeTableName;
                pasteWorkDataModel.Update_end_time = copiedWorkDataModel.Update_end_time;
                pasteWorkDataModel.Update_start_time = copiedWorkDataModel.Update_start_time;
                pasteWorkDataModel.Work_days = copiedWorkDataModel.Work_days;
                pasteWorkDataModel.Work_from = copiedWorkDataModel.Work_from;
                pasteWorkDataModel.Work_to = copiedWorkDataModel.Work_to;
                pasteWorkDataModel.Work_type_no = copiedWorkDataModel.Work_type_no;
                pasteWorkDataModel.Working_time = copiedWorkDataModel.Working_time;
                isCopyMode = false;
            }
        }

        protected override void OnDeactivate(bool close)
        {
            base.OnDeactivate(close);
        }
        protected override void OnActivate()
        {
            base.OnActivate();
            decimal firstDayOfMonth = CommonUtil.GetFirstDayOfMonth(currentYearMonth);
            decimal lastDayOfMonth = CommonUtil.GetLastDayOfMonth(currentYearMonth);
            timeTableList = timeTableService.SearchTimeTableList();
            holidayList = holidayService.SearchHolidayList(companyNo, firstDayOfMonth, lastDayOfMonth);
            GetData();
        }
        public override void CanClose(Action<bool> callback)
        {
            if (isChanged)
            {
                var result = MessageBox.Show(ResourcesManager.GetMessage(MessageConstant.I0002).Message, MessageConstant.I0002, MessageBoxButton.YesNo, MessageBoxImage.Information);
                if (result == MessageBoxResult.No)
                {
                    return;
                }
                workDataService.UpdateWorkDataList(this.workDataList);
            }
            else if (hasAutoChangeContactTime)
            {
                workDataService.UpdateWorkDataList(this.workDataList);
            }
            base.CanClose(callback);
        }
        protected override void SetMultiLanguage()
        {
            base.SetMultiLanguage();
            LanguageModel form = ResourcesManager.GetLanguageForForm(FORM_ID);
            var display = ResourcesManager.GetLanguageForControlInForm(form, "Text");
            if (display != null)
            {
                SetDisplayName(display.Text);
            }

            LblAttendenceTitle = ResourcesManager.GetLanguageForControlInForm(form, "Title");
            LblAttendenceName = ResourcesManager.GetLanguageForControlInForm(form, "Label1");
            LblMessageArea = ResourcesManager.GetLanguageForControlInForm(form, "MessageArea");
            LblMessageText = ResourcesManager.GetLanguageForControlInForm(form, "MessageText");
            LblEmployeeName = ResourcesManager.GetLanguageForControlInForm(form, "lblEmployeeName");

            LblAttendencePeriod = ResourcesManager.GetLanguageForControlInForm(form, "Label3");
            LblPrev = ResourcesManager.GetLanguageForControlInForm(form, "btnPreviousMonth");


            LblPeriod = ResourcesManager.GetLanguageForControlInForm(form, "lblPeriodFromTo");
            var tag = ResourcesManager.GetLanguageForControlInForm(LblPeriod, "Tag");
            if (tag != null)
            {
                dateFormat = tag.Text;
            }
            if (string.IsNullOrEmpty(dateFormat))
            {
                dateFormat = "YMD";
            }
            LblNext = ResourcesManager.GetLanguageForControlInForm(form, "btnNextMonth");
            LblPersonalExport = ResourcesManager.GetLanguageForControlInForm(form, "btnCustomizeList1XlsExport");
            var dataGridView = ResourcesManager.GetLanguageForControlInForm(form, "DataGridView");
            var dataGridHeader = ResourcesManager.GetLanguageForControlInForm(dataGridView, "HeaderText");
            if (dataGridHeader != null && !string.IsNullOrEmpty(dataGridHeader.Text))
            {
                HeaderText = dataGridHeader.Text.Split(',');
            }
            var workDateClm = ResourcesManager.GetLanguageForControlInForm(dataGridView, "Tag");
            if (workDateClm != null)
            {
                workDateFormat = workDateClm.Text;
            }
            LblCancel = ResourcesManager.GetLanguageForControlInForm(form, "btnPrevious");
            LblUpdate = ResourcesManager.GetLanguageForControlInForm(form, "btnUpdate");
            var work_type_noColumn = ResourcesManager.GetLanguageForControlInForm(form, "WORK_TYPE_NOColumn");
            List<SelectItemModel> dayTypeList = new List<SelectItemModel>();
            if (work_type_noColumn != null)
            {
                DayTypeList = CommonUtil.CreateSelectItemList(work_type_noColumn.Text);
            }
            else
            {
                DayTypeList = new List<SelectItemModel>();
            }
        }

        private bool DoValidate()
        {
            if (workDataList == null)
            {
                return false;
            }
            bool isValid = true;


            foreach (var work in workDataList)
            {

                isValid = DoCheckInput(work);
                if (!isValid)
                {
                    break;
                }
            }

            return isValid;
        }

        public void Save()
        {
            if (!DoValidate())
            {
                return;
            }
            if (isChanged)
            {
                var result = MessageBox.Show(ResourcesManager.GetMessage(MessageConstant.I0001).Message, MessageConstant.I0001, MessageBoxButton.YesNo, MessageBoxImage.Information);
                if (result == MessageBoxResult.No)
                {
                    return;
                }
                workDataService.UpdateWorkDataList(this.workDataList);
            }
            else if (hasAutoChangeContactTime)
            {
                workDataService.UpdateWorkDataList(this.workDataList);
            }
            isChanged = false;
            TryClose(true);
        }

        public void Prev()
        {
            if (!DoValidate())
            {
                return;
            }

            decimal lastMonth = CommonUtil.GetLastMonth(currentYearMonth);
            decimal firstDayOfMonth = CommonUtil.GetFirstDayOfMonth(lastMonth);
            decimal lastDayOfMonth = CommonUtil.GetLastDayOfMonth(lastMonth);
            if (!(IsEmployeeInExpiration(firstDayOfMonth) || IsEmployeeInExpiration(lastDayOfMonth)))
            {
                Message = ResourcesManager.GetMessage(MessageConstant.A0027);
                return;
            }

            if (isChanged)
            {
                var result = MessageBox.Show(ResourcesManager.GetMessage(MessageConstant.I0006).Message, MessageConstant.I0006, MessageBoxButton.YesNoCancel, MessageBoxImage.Information);
                if (result == MessageBoxResult.Cancel)
                {
                    return;
                }
                else if (result == MessageBoxResult.Yes)
                {
                    workDataService.UpdateWorkDataList(this.workDataList);
                }
            }
            else if (hasAutoChangeContactTime)
            {
                workDataService.UpdateWorkDataList(this.workDataList);
            }
            currentYearMonth = lastMonth;
            holidayList = holidayService.SearchHolidayList(companyNo, firstDayOfMonth, lastDayOfMonth);
            GetData();
        }

        public void Next()
        {
            if (!DoValidate())
            {
                return;
            }
            decimal nextMonth = CommonUtil.GetNextMonth(currentYearMonth);
            decimal firstDayOfMonth = CommonUtil.GetFirstDayOfMonth(nextMonth);
            decimal lastDayOfMonth = CommonUtil.GetLastDayOfMonth(nextMonth);
            if (!(IsEmployeeInExpiration(firstDayOfMonth) || IsEmployeeInExpiration(lastDayOfMonth)))
            {
                Message = ResourcesManager.GetMessage(MessageConstant.A0027);
                return;
            }

            if (isChanged)
            {

                var result = MessageBox.Show(ResourcesManager.GetMessage(MessageConstant.I0007).Message, MessageConstant.I0007, MessageBoxButton.YesNoCancel, MessageBoxImage.Information);
                if (result == MessageBoxResult.Cancel)
                {
                    return;
                }
                else if (result == MessageBoxResult.Yes)
                {
                    workDataService.UpdateWorkDataList(this.workDataList);
                }
            }
            else if (hasAutoChangeContactTime)
            {
                workDataService.UpdateWorkDataList(this.workDataList);
            }

            currentYearMonth = nextMonth;

            holidayList = holidayService.SearchHolidayList(companyNo, firstDayOfMonth, lastDayOfMonth);

            GetData();

        }

        public void Cancel()
        {
            if (isChanged)
            {
                var result = MessageBox.Show(ResourcesManager.GetMessage(MessageConstant.I0002).Message, MessageConstant.I0002, MessageBoxButton.YesNo, MessageBoxImage.Information);
                if (result == MessageBoxResult.No)
                {
                    return;
                }
                workDataService.UpdateWorkDataList(this.workDataList);
                isChanged = false;
            }
            else if (hasAutoChangeContactTime)
            {
                workDataService.UpdateWorkDataList(this.workDataList);
            }
            this.TryClose(true);
        }

        public void PersonalExport()
        {
            if (!DoValidate())
            {
                return;
            }
            if (isChanged)
            {
                var result = MessageBox.Show(ResourcesManager.GetMessage(MessageConstant.I0001).Message, MessageConstant.I0001, MessageBoxButton.YesNo, MessageBoxImage.Information);
                if (result == MessageBoxResult.Yes)
                {
                    workDataService.UpdateWorkDataList(this.workDataList);
                }
                else
                {
                    return;
                }

            }
            string templateFile = Properties.Settings.Default.XLS_Personal_Report;
            if (Properties.Settings.Default.XLS_Use_Multi_Language)
            {
                templateFile = string.Format(templateFile, Properties.Settings.Default.SelectedLanguage);
            }
            else
            {
                templateFile = string.Format(templateFile, string.Empty);
            }

            string outFileName = string.Format(Properties.Settings.Default.XLS_Out_Personal_File, employeeNo, CommonUtil.GetDateAsString(currentYearMonth, "MY").ToUpper() + ".");

            Export(outFileName, templateFile);

        }

        private bool Export(string outFileName, string templateFileName)
        {
            bool success = false;

            try
            {
                CreateExcelApplication();

                tempWorkBook = excelApp.Workbooks.Open(CommonUtil.GetTemplate(templateFileName), 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
                workBook = CommonUtil.CreateWorkbook(excelApp, outFileName);
                Excel.Worksheet tempSheet = tempWorkBook.Worksheets["Employee_no"];
                tempSheet.Copy(workBook.Worksheets[Properties.Settings.Default.XLS_Default_Sheet_Name], Type.Missing);
                CommonUtil.DeleteDefaultSheet(workBook);

                Excel.Worksheet employeeSheet = (Microsoft.Office.Interop.Excel.Worksheet)workBook.Worksheets["Employee_no"];
                CommonUtil.SetXlsPageSetup(employeeSheet, XlPageOrientation.xlLandscape);
                CommonUtil.FillPersonalReportSheet(employeeSheet, employeeModel, workDataList, dayTypeList, timeTableList, unit_minutes);
                tempWorkBook.Saved = true;
                tempWorkBook.Close();
                Show();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Exception", MessageBoxButton.OK, MessageBoxImage.Error);
                excelApp.Quit();
            }
            finally
            {
                DisposeXls();
            }
            return success;
        }

        private void DisposeXls()
        {

            if (tempWorkBook != null)
            {
                Marshal.ReleaseComObject(RuntimeHelpers.GetObjectValue(this.tempWorkBook));
            }
            if (this.workBook != null)
            {
                Marshal.ReleaseComObject(RuntimeHelpers.GetObjectValue(this.workBook));
            }

            if (this.excelApp != null)
            {
                Marshal.ReleaseComObject(RuntimeHelpers.GetObjectValue(this.excelApp));
            }
            GC.Collect();
        }

        private void CreateExcelApplication()
        {

            excelApp = (Excel.Application)RuntimeHelpers.GetObjectValue(new Excel.Application());
            excelApp.DisplayAlerts = false;
            excelApp.Visible = false;
        }

        private void Show()
        {
            this.excelApp.DisplayAlerts = true;
            this.excelApp.WindowState = XlWindowState.xlMaximized;
            this.excelApp.Visible = true;
            this.excelApp.SendKeys("{ESC}");
            foreach (Process process in Process.GetProcessesByName("Excel"))
            {
                if (process.MainWindowTitle.TrimEnd(new char[0]).EndsWith(((Workbook)this.workBook).FullName))
                {
                    CommonUtil.SetWindowPosision(0L, process.Handle);
                    return;
                }
            }
        }

        #region get/set
        public List<WorkDataModel> WorkDataList
        {
            get
            {
                return workDataList;
            }
            set
            {
                if (workDataList != value)
                {
                    workDataList = value;
                    NotifyOfPropertyChange(() => WorkDataList);
                }
            }
        }

        public WorkDataModel SelectedWorkData
        {
            get
            {
                return selectedWorkData;
            }
            set
            {
                if (selectedWorkData != value)
                {
                    selectedWorkData = value;
                    NotifyOfPropertyChange(() => SelectedWorkData);
                }
            }
        }
        public WorkDataModel CopiedWorkDataModel
        {
            get
            {
                return copiedWorkDataModel;
            }
            set
            {
                if (copiedWorkDataModel != value)
                {
                    copiedWorkDataModel = value;
                    NotifyOfPropertyChange(() => CopiedWorkDataModel);
                }
            }
        }

        public WorkDataModel PasteWorkDataModel
        {
            get
            {
                return pasteWorkDataModel;
            }
            set
            {
                if (pasteWorkDataModel != value)
                {
                    pasteWorkDataModel = value;
                    NotifyOfPropertyChange(() => PasteWorkDataModel);
                }
            }
        }

        public LanguageModel LblAttendenceTitle
        {
            get
            {
                return lblAttendenceTitle;
            }
            set
            {
                if (lblAttendenceTitle != value)
                {
                    lblAttendenceTitle = value;
                    NotifyOfPropertyChange(() => LblAttendenceTitle);
                }
            }
        }

        public LanguageModel LblAttendenceName
        {
            get
            {
                return lblAttendenceName;
            }
            set
            {
                if (lblAttendenceName != value)
                {
                    lblAttendenceName = value;
                    NotifyOfPropertyChange(() => LblAttendenceName);
                }
            }
        }

        public LanguageModel LblAttendencePeriod
        {
            get
            {
                return lblAttendencePeriod;
            }
            set
            {
                if (lblAttendencePeriod != value)
                {
                    lblAttendencePeriod = value;
                    NotifyOfPropertyChange(() => LblAttendencePeriod);
                }
            }
        }

        public LanguageModel LblPrev
        {
            get
            {
                return lblPrev;
            }
            set
            {
                if (lblPrev != value)
                {
                    lblPrev = value;
                    NotifyOfPropertyChange(() => LblPrev);
                }
            }
        }

        public LanguageModel LblNext
        {
            get
            {
                return lblNext;
            }
            set
            {
                if (lblNext != value)
                {
                    lblNext = value;
                    NotifyOfPropertyChange(() => LblNext);
                }
            }
        }

        public LanguageModel LblMessageArea
        {
            get
            {
                return lblMessageArea;
            }
            set
            {
                if (lblMessageArea != value)
                {
                    lblMessageArea = value;
                    NotifyOfPropertyChange(() => LblMessageArea);
                }
            }
        }

        public LanguageModel LblMessageText
        {
            get
            {
                return lblMessageText;
            }
            set
            {
                if (lblMessageText != value)
                {
                    lblMessageText = value;
                    NotifyOfPropertyChange(() => LblMessageText);
                }
            }
        }

        public LanguageModel LblEmployeeName
        {
            get
            {
                return lblEmployeeName;
            }
            set
            {
                if (lblEmployeeName != value)
                {
                    lblEmployeeName = value;
                    NotifyOfPropertyChange(() => LblEmployeeName);
                }
            }
        }

        public LanguageModel LblPeriod
        {
            get
            {
                return lblPeriod;
            }
            set
            {
                if (lblPeriod != value)
                {
                    lblPeriod = value;
                    NotifyOfPropertyChange(() => LblPeriod);
                }
            }
        }

        public LanguageModel LblPersonalExport
        {
            get
            {
                return lblPersonalExport;
            }
            set
            {
                if (lblPersonalExport != value)
                {
                    lblPersonalExport = value;
                    NotifyOfPropertyChange(() => LblPersonalExport);
                }
            }
        }

        public LanguageModel LblCancel
        {
            get
            {
                return lblCancel;
            }
            set
            {
                if (lblCancel != value)
                {
                    lblCancel = value;
                    NotifyOfPropertyChange(() => LblCancel);
                }
            }
        }

        public string[] HeaderText
        {
            get
            {
                return headerText;
            }
            set
            {
                if (headerText != value)
                {
                    headerText = value;
                    NotifyOfPropertyChange(() => HeaderText);
                }
            }
        }

        public bool IsFocused
        {
            get
            {
                return isFocused;
            }
            set
            {
                if (isFocused != value)
                {
                    isFocused = value;
                    NotifyOfPropertyChange(() => IsFocused);
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


        public LanguageModel LblUpdate
        {
            get
            {
                return lblUpdate;
            }
            set
            {
                if (lblUpdate != value)
                {
                    lblUpdate = value;
                    NotifyOfPropertyChange(() => LblUpdate);
                }
            }
        }
        public string Period
        {
            get
            {
                return period;
            }
            set
            {
                if (period != value)
                {
                    period = value;
                    NotifyOfPropertyChange(() => Period);
                }
            }
        }
        public List<SelectItemModel> DayTypeList
        {
            get
            {
                return dayTypeList;
            }
            set
            {
                if (dayTypeList != value)
                {
                    dayTypeList = value;
                    NotifyOfPropertyChange(() => DayTypeList);
                }
            }
        }
        public List<TimeTableModel> TimeTableList
        {
            get
            {
                return timeTableList;
            }
            set
            {
                if (timeTableList != value)
                {
                    timeTableList = value;
                    NotifyOfPropertyChange(() => TimeTableList);
                }
            }
        }

        public string DeletePropertyName
        {
            get
            {
                return deletePropertyName;
            }
            set
            {
                if (deletePropertyName != value)
                {
                    deletePropertyName = value;
                    NotifyOfPropertyChange(() => DeletePropertyName);
                }
            }
        }

        #endregion
    }
}
