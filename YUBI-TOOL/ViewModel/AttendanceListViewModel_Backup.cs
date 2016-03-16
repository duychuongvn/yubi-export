using System.Collections.Generic;
using System.Linq;
using System.Windows;
using Caliburn.Micro;
using YUBI_TARO.EXP.Common;
using YUBI_TARO.EXP.Model;
using YUBI_TARO.EXP.Service;
using System;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using Microsoft.Win32;

namespace YUBI_TARO.EXP.ViewModel
{
    public class AttendanceListViewModel : ViewModelBase
    {
        private const string FORM_ID = "WorkList";

        private readonly IWorkDataService workDataService;
        private readonly IEmployeeService employeeService;
        private readonly ITimeTableService timeTableService;
        private readonly IHolidayService holidayService;
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
        private WorkDataModel workDataCopied;
        private bool isChanged = false;
        public AttendanceListViewModel()
        {
            workDataService = IoC.Get<IWorkDataService>();
            employeeService = IoC.Get<IEmployeeService>();
            holidayService = IoC.Get<IHolidayService>();
            timeTableService = IoC.Get<ITimeTableService>();
        }
        public void Init(string employeeNo, decimal companyNo, decimal postNo, decimal yearMonth)
        {
            this.employeeNo = employeeNo;
            this.companyNo = companyNo;
            this.post_no = postNo;
            this.currentYearMonth = yearMonth;
            this.employeeModel = employeeService.GetEmployee(companyNo, post_no, employeeNo, yearMonth);
            if (employeeModel != null)
            {
                EmployeeName = CommonUtil.GetFullName(employeeModel.Emsize_first_name, employeeModel.Emsize_last_name);
            }
        }

        private void GetData()
        {
            decimal firstDayOfMonth = CommonUtil.GetFirstDayOfMonth(currentYearMonth);
            decimal lastDayOfMonth = CommonUtil.GetLastDayOfMonth(currentYearMonth);
            Period = CommonUtil.GetDateAsString(firstDayOfMonth, dateFormat, false) + " - " + CommonUtil.GetDateAsString(CommonUtil.GetLastDayOfMonth(currentYearMonth), dateFormat, false);
            WorkDataList = workDataService.SearchWorkDataListByEmployee(companyNo, post_no, employeeNo, firstDayOfMonth, lastDayOfMonth);
            ParseData();
            isChanged = false;

        }

        private void ParseData()
        {
            foreach (var workData in workDataList)
            {
                workData.Work_date_dsp = CommonUtil.GetDateAsString(workData.Work_date, workDateFormat, true);
                if (employeeModel.Use_flag_of_holiday == DBConstant.FLAG_OF_HOLIDAY_USE && holidayList.Count(x => x.Holiday_date == workData.Work_date) > 0)
                {
                    workData.IsHoliday = true;
                }
                if (workData.Work_day_type_no != DBConstant.WORK_DAY_TYPE_NORMAL && workData.Work_day_type_no != DBConstant.WORK_DAY_TYPE_REGULAR)
                {
                    workData.IsHoliday = true;
                }
                var dayType = dayTypeList.Find(x => CommonUtil.ToDecimal(x.ItemCD) == workData.Work_type_no);
                var timeTable = timeTableList.Find(x => x.Time_table_no == workData.Time_table_no);
                if (dayType != null)
                {
                    workData.Work_type_name = dayType.ItemValue;
                }
                if (timeTable != null)
                {
                    workData.TimeTableName = timeTable.Time_table_name;
                }
                workData.PropertyChanged -= new System.ComponentModel.PropertyChangedEventHandler(workData_PropertyChanged);
                workData.PropertyChanged += new System.ComponentModel.PropertyChangedEventHandler(workData_PropertyChanged);
            }
        }


        private void workData_PropertyChanged(object sender, System.ComponentModel.PropertyChangedEventArgs e)
        {
            WorkDataModel workData = sender as WorkDataModel;
            isChanged = true;
            if (e.PropertyName == "Time_table_no")
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

            }
            else if (e.PropertyName == "Update_start_time"
               || e.PropertyName == "Update_end_time"
               || e.PropertyName == "Work_type_no")
            {
                if (workData.Work_type_no == DBConstant.WORK_DAY_TYPE_NORMAL
                    || workData.Work_type_no == DBConstant.WORK_TYPE_HOLIDAY_DUTY
                    || workData.Work_type_no == DBConstant.WORK_TYPE_HALF_PERMISSION)
                {
                    var timeTable = timeTableList.Find(x => x.Time_table_no == workData.Time_table_no);
                    decimal? update_start_time = CommonUtil.ToNullableDecimal(workData.Update_start_time);
                    decimal? update_end_time = CommonUtil.ToNullableDecimal(workData.Update_end_time);
                    if (update_start_time != null && update_end_time != null)
                    {

                        string restFrom = "Rest{0}_from";
                        string restTo = "Rest{0}_to";
                        decimal restTime = 0;
                        decimal overTime = 0;
                        decimal lateness = 0;
                        decimal leaveEarly = 0;

                        if (timeTable.Work_from < update_start_time)
                        {
                            lateness = CommonUtil.SubTime(timeTable.Work_from, update_start_time.Value);
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
                                    else
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
                        if (lateness > 0)
                        {
                            workData.Being_late_time = CommonUtil.ToString(lateness);
                        }
                        workData.Rest_time = CommonUtil.ToString(restTime);
                        workData.Over_time = CommonUtil.ToString(overTime);
                        decimal working_time = CommonUtil.SubTime(timeTable.Work_to, timeTable.Work_from) - leaveEarly - lateness + overTime;
                        if (workData.Work_type_no == DBConstant.WORK_TYPE_HOLIDAY_DUTY)
                        {
                            workData.Working_time = null;
                            workData.Holiday_time = CommonUtil.ToString(working_time);
                        }
                        else
                        {
                            workData.Holiday_time = null;
                            workData.Working_time = CommonUtil.ToString(working_time);
                        }

                    }
                    else
                    {

                        workData.Working_time = null;
                        workData.Rest_time = null;
                        workData.Over_time = null;
                        workData.Being_late_time = null;
                        workData.Leaving_early_time = null;
                    }
                }
                else
                {

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

        public void DoCopy(object workData)
        {
            //workDataCopied = workData;
        }
        public void DoPaste(WorkDataModel workData)
        {
            if (workData != null && workDataCopied != null)
            {
                workData.Update_start_time = workDataCopied.Update_start_time;
                workData.Update_end_time = workDataCopied.Update_end_time;
                workData.Working_time = workDataCopied.Working_time;
            }
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
        public void Save()
        {
            if (isChanged)
            {
                var result = MessageBox.Show(ResourcesManager.GetMessage(MessageConstant.I0001).Message, MessageConstant.I0001, MessageBoxButton.YesNo, MessageBoxImage.Information);
                if (result == MessageBoxResult.No)
                {
                    return;
                }
                workDataService.UpdateWorkDataList(this.workDataList);
            }
            TryClose(true);
        }
        public void Prev()
        {
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
            currentYearMonth = CommonUtil.GetLastMonth(currentYearMonth);
            decimal firstDayOfMonth = CommonUtil.GetFirstDayOfMonth(currentYearMonth);
            decimal lastDayOfMonth = CommonUtil.GetLastDayOfMonth(currentYearMonth);
            holidayList = holidayService.SearchHolidayList(companyNo, firstDayOfMonth, lastDayOfMonth);
            GetData();
        }
        public void Next()
        {
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
            currentYearMonth = CommonUtil.GetNextMonth(currentYearMonth);
            decimal firstDayOfMonth = CommonUtil.GetFirstDayOfMonth(currentYearMonth);
            decimal lastDayOfMonth = CommonUtil.GetLastDayOfMonth(currentYearMonth);
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
            }
            this.TryClose(true);
        }

        public void PersonalExport()
        {
            if (isChanged)
            {
                var result = MessageBox.Show(ResourcesManager.GetMessage(MessageConstant.I0001).Message, MessageConstant.I0001, MessageBoxButton.YesNo, MessageBoxImage.Information);
                if (result == MessageBoxResult.Yes)
                {
                    workDataService.UpdateWorkDataList(this.workDataList);
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
            string outFileName = string.Format(Properties.Settings.Default.XLS_Out_Personal_file, employeeNo, CommonUtil.GetDateAsString(currentYearMonth, "MY"));
            if (string.IsNullOrEmpty(Properties.Settings.Default.XLS_Export_Path)
                || !Directory.Exists(Properties.Settings.Default.XLS_Export_Path))
            {
                SaveFileDialog saveFile = new SaveFileDialog();
                saveFile.Filter = "Microsoft Office Excel Workbook(*.xls)|*.xls";
                saveFile.FileName = outFileName;
                bool? result = saveFile.ShowDialog();
                if (result == null || !result.Value)
                {
                    return;
                }
                Properties.Settings.Default.XLS_Export_Path = saveFile.FileName.Replace(saveFile.SafeFileName, string.Empty);
                Properties.Settings.Default.Save();
                outFileName = saveFile.FileName;
            }
            else
            {
                outFileName = Properties.Settings.Default.XLS_Export_Path + outFileName;
            }

            bool sucess = Export(outFileName, templateFile);
            if (sucess)
            {
                //System.Diagnostics.Process.Start(outFileName);
            }
        }

        private bool Export(string outFileName, string templateFileName)
        {
            bool success = false;
            templateFileName = Environment.CurrentDirectory + @"\Template\" + templateFileName;
            Microsoft.Office.Interop.Excel.Application excelApp = null;
            try
            {

                excelApp = new Microsoft.Office.Interop.Excel.Application();
                // bool a =  excelApp.NewWorkbook.Add(outFileName, Type.Missing, "ChuongTest", Type.Missing);
                //Excel.Workbook employeeWorkBook = excelApp.Workbooks.Add(Type.Missing);
                Excel.Workbook employeeWorkBook = excelApp.Workbooks.Add(templateFileName);
                //  Excel.Workbook tempWorkBook = excelApp.Workbooks.Open(templateFileName, 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
                //Excel.Worksheet tempSheet = tempWorkBook.Worksheets["Employee_no"];
                // tempSheet.Copy(employeeWorkBook.Worksheets["Sheet1"], Type.Missing);
                //employeeWorkBook.Worksheets["Sheet1"].Delete();
                //employeeWorkBook.Worksheets["Sheet2"].Delete();
                //employeeWorkBook.Worksheets["Sheet3"].Delete();
                Microsoft.Office.Interop.Excel.Worksheet employeeSheet = (Microsoft.Office.Interop.Excel.Worksheet)employeeWorkBook.Worksheets["Employee_no"];
                employeeSheet.Name = employeeNo;

                //  tempWorkBook.Close();

                Excel.Range employeeNameRange = employeeSheet.get_Range("G6", "H6");
                Excel.Range employeeCodeRange = employeeSheet.get_Range("G7", "H7");
                Excel.Range departmentRange = employeeSheet.get_Range("L6");
                Excel.Range positionRange = employeeSheet.get_Range("L7");
                Excel.Range teamRange = employeeSheet.get_Range("L8");
                Excel.Range detailRange = employeeSheet.get_Range("A11", "S11");

                int startRow = 11;
                string detailRowStart = "A{0}";
                string detailRowEnd = "S{0}";
                decimal totalWorkingTime = 0;
                decimal totalOverTime = 0;
                decimal totalHolidayTime = 0;
                decimal totalMidleNight = 0;
                decimal totalHolidayMidleNigh = 0;
                decimal totalPublicHolidayTime = 0;
                decimal totalPublicMidleNightHolidayTime = 0;
                decimal totalMidleNightHolidayTime = 0;
                decimal contactTime = 0;
                int totalAL = 0;
                int totalP = 0;
                int totalWP = 0;
                int totalSL = 0;
                employeeNameRange.Value = employeeName;
                employeeCodeRange.Value = employeeNo;
                departmentRange.Value = employeeModel.Post_name;
                if (!string.IsNullOrEmpty(employeeModel.Remarks) && employeeModel.Remarks.Length > DBConstant.EMPL_TEAM_LEN)
                {
                    positionRange.Value = employeeModel.Remarks.Substring(DBConstant.EMPL_TEAM_LEN).Trim();
                    teamRange.Value = employeeModel.Remarks.Substring(0, DBConstant.EMPL_TEAM_LEN).Trim();
                }
                else if (employeeModel.Remarks != null)
                {
                    teamRange.Value = employeeModel.Remarks.Trim();
                }

                foreach (var workData in workDataList)
                {
                    int columnIndex = 1;
                    int aL = 0;
                    int p = 0;
                    int wp = 0;
                    int sl = 0;

                    TimeTableModel timeTable = timeTableList.Find(x => x.Time_table_no == workData.Time_table_no);
                    var work_type = dayTypeList.Find(x => decimal.Parse(x.ItemCD) == workData.Work_type_no);
                    string time = "";
                    string work_type_name = "";
                    if (timeTable != null)
                    {
                        time = CommonUtil.ToDispHour(timeTable.Work_from) + "-" + CommonUtil.ToDispHour(timeTable.Work_to);
                        if (!workData.IsHoliday && (workData.Work_type_no == DBConstant.WORK_TYPE_NORMAL || workData.Work_type_no == DBConstant.WORK_TYPE_HALF_PERMISSION))
                        {
                            if (workData.Work_type_no == DBConstant.WORK_TYPE_HALF_PERMISSION)
                            {
                                contactTime += CommonUtil.SubTime(timeTable.Work_to, timeTable.Work_from) / 2;
                            }
                            else
                            {
                                contactTime += CommonUtil.SubTime(timeTable.Work_to, timeTable.Work_from);
                            }

                        }
                    }
                    if (work_type != null)
                    {
                        work_type_name = work_type.ItemValue;
                    }
                    decimal? workingTime = CommonUtil.ToNullableDecimal(workData.Working_time);
                    decimal? overTime = CommonUtil.ToNullableDecimal(workData.Over_time);
                    decimal? hoildayTime = null;
                    decimal? pubHoildayTime = null;
                    decimal? midleNightHoildayTime = null;
                    decimal? pupMidleNightHoildayTime = null;
                    decimal? midleNight = null;
                    if (workData.IsHoliday)
                    {
                        if (workData.Work_day_type_no != DBConstant.WORK_DAY_TYPE_REGULAR && workData.Work_day_type_no != DBConstant.WORK_DAY_TYPE_NORMAL)
                        {
                            pubHoildayTime = CommonUtil.ToNullableDecimal(workData.Holiday_time);
                            pupMidleNightHoildayTime = CommonUtil.ToNullableDecimal(workData.Holiday_late_night_time);
                            if (pubHoildayTime != null)
                            {
                                totalPublicHolidayTime += pubHoildayTime.Value;
                            }
                            if (pupMidleNightHoildayTime != null)
                            {
                                totalPublicMidleNightHolidayTime += pupMidleNightHoildayTime.Value;
                            }
                        }
                        else
                        {
                            hoildayTime = CommonUtil.ToNullableDecimal(workData.Holiday_time);
                            midleNightHoildayTime = CommonUtil.ToNullableDecimal(workData.Holiday_late_night_time);
                            if (hoildayTime != null)
                            {
                                totalHolidayTime += hoildayTime.Value;
                            }

                            if (midleNightHoildayTime != null)
                            {
                                totalMidleNightHolidayTime += midleNightHoildayTime.Value;
                            }
                        }
                    }
                    else
                    {
                        midleNight = CommonUtil.ToNullableDecimal(workData.Late_night_time);
                        if (midleNight != null)
                        {
                            totalMidleNight += midleNight.Value;
                        }
                    }

                    if (workingTime != null)
                    {
                        totalWorkingTime += workingTime.Value;
                    }
                    if (overTime != null)
                    {
                        totalOverTime += overTime.Value;
                    }
                    switch ((int)workData.Work_type_no.Value)
                    {
                        case DBConstant.WORK_TYPE_ANNUAL_LEAVE:
                            aL = 1;
                            totalAL++;
                            break;
                        case DBConstant.WORK_TYPE_PERMISSION:
                            p = 1;
                            totalP++;
                            break;
                        case DBConstant.WORK_TYPE_AW_PERMISSION:
                            wp = 1;
                            totalWP++;
                            break;
                        case DBConstant.WORK_TYPE_SPECIAL_LEAVE:
                            sl = 1;
                            totalSL++;
                            break;
                        default:
                            break;
                    }

                    detailRange.Insert(Type.Missing, Type.Missing);
                    Excel.Range detailEditRange = employeeSheet.get_Range(string.Format(detailRowStart, startRow), string.Format(detailRowEnd, startRow));
                    detailEditRange.Columns[columnIndex++].Value = CommonUtil.FindDate(workData.Work_date_dsp);
                    detailEditRange.Columns[columnIndex++].Value = CommonUtil.FindWeekDay(workData.Work_date_dsp);
                    detailEditRange.Columns[columnIndex++].Value = work_type_name;
                    detailEditRange.Columns[columnIndex++].Value = time;
                    detailEditRange.Columns[columnIndex++].Value = CommonUtil.ToDispHour(workData.Start_time);
                    detailEditRange.Columns[columnIndex++].Value = CommonUtil.ToDispHour(workData.End_time);
                    detailEditRange.Columns[columnIndex++].Value = CommonUtil.ToDispMinute(workData.Working_time);
                    detailEditRange.Columns[columnIndex++].Value = CommonUtil.ToDispMinute(workData.Over_time);
                    detailEditRange.Columns[columnIndex++].Value = CommonUtil.ToDispMinute(workData.Late_night_time);
                    detailEditRange.Columns[columnIndex++].Value = CommonUtil.ToDispMinute(workData.Holiday_time);
                    detailEditRange.Columns[columnIndex++].Value = CommonUtil.ToDispMinute(workData.Holiday_late_night_time);

                    detailEditRange.Columns[columnIndex++].Value = (aL != 0) ? aL.ToString() : "";
                    detailEditRange.Columns[columnIndex++].Value = (p != 0) ? p.ToString() : "";
                    detailEditRange.Columns[columnIndex++].Value = (wp != 0) ? wp.ToString() : "";
                    detailEditRange.Columns[columnIndex++].Value = (sl != 0) ? sl.ToString() : "";
                    detailEditRange.Columns[columnIndex++].Value = workData.Memo;
                    startRow++;
                }
                startRow += 1;
                Excel.Range footer0Range = employeeSheet.get_Range("D" + startRow, "O" + startRow++);
                Excel.Range footer1Range = employeeSheet.get_Range(string.Format(detailRowStart, startRow), string.Format(detailRowEnd, startRow++));
                Excel.Range footer2Range = employeeSheet.get_Range(string.Format(detailRowStart, startRow), string.Format(detailRowEnd, startRow++));
                int col_contact_Ot = 4;
                int col_working_Mid = 7;
                int col_Holiday_H_Mid = 11;
                int col_P_Holiday_P_Mid = 15;
                int col_total_start_index = 4;
                footer0Range.Columns[col_total_start_index++].Value = CommonUtil.ToDispMinute(totalWorkingTime);
                footer0Range.Columns[col_total_start_index++].Value = CommonUtil.ToDispMinute(totalOverTime);
                footer0Range.Columns[col_total_start_index++].Value = CommonUtil.ToDispMinute(totalMidleNight);
                footer0Range.Columns[col_total_start_index++].Value = CommonUtil.ToDispMinute(totalHolidayTime + totalPublicHolidayTime);
                footer0Range.Columns[col_total_start_index++].Value = CommonUtil.ToDispMinute(totalHolidayMidleNigh);
                footer0Range.Columns[col_total_start_index++].Value = totalAL > 0 ? totalAL.ToString() : "";
                footer0Range.Columns[col_total_start_index++].Value = totalP > 0 ? totalP.ToString() : "";
                footer0Range.Columns[col_total_start_index++].Value = totalWP > 0 ? totalWP.ToString() : "";
                footer0Range.Columns[col_total_start_index++].Value = totalSL > 0 ? totalSL.ToString() : "";


                footer1Range.Columns[col_contact_Ot].Value = CommonUtil.ToDispMinute(contactTime);
                footer1Range.Columns[col_working_Mid].Value = CommonUtil.ToDispMinute(totalWorkingTime + totalHolidayTime + totalPublicHolidayTime);
                footer1Range.Columns[col_Holiday_H_Mid].Value = CommonUtil.ToDispMinute(totalHolidayTime);
                footer1Range.Columns[col_P_Holiday_P_Mid].Value = CommonUtil.ToDispMinute(totalPublicHolidayTime);

                footer2Range.Columns[col_contact_Ot].Value = CommonUtil.ToDispMinute(totalOverTime);
                footer2Range.Columns[col_working_Mid].Value = CommonUtil.ToDispMinute(totalMidleNight);
                footer1Range.Columns[col_Holiday_H_Mid].Value = CommonUtil.ToDispMinute(totalHolidayMidleNigh);
                footer2Range.Columns[col_P_Holiday_P_Mid].Value = CommonUtil.ToDispMinute(totalPublicMidleNightHolidayTime);

                // employeeWorkBook.Name = outFileName;
                //excelApp.SaveWorkspace("ChuongTest");

                excelApp.Application.Visible = true;
                employeeWorkBook.Saved = false;
                //employeeWorkBook.SaveAs(outFileName, Microsoft.Office.Interop.Excel.XlFileFormat.xlExcel9795,
                //    System.Reflection.Missing.Value, System.Reflection.Missing.Value, false, false,
                //    Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange,
                //    Microsoft.Office.Interop.Excel.XlSaveConflictResolution.xlUserResolution, true,
                //    System.Reflection.Missing.Value,
                //    System.Reflection.Missing.Value,
                //    System.Reflection.Missing.Value);
                //employeeWorkBook.Saved = true;
                success = true;
               // employeeWorkBook.BeforeSave += new Excel.WorkbookEvents_BeforeSaveEventHandler(employeeWorkBook_BeforeSave);
                excelApp.WorkbookBeforeSave += new Excel.AppEvents_WorkbookBeforeSaveEventHandler(excelApp_WorkbookBeforeSave);
                excelApp.WorkbookBeforeClose += new Excel.AppEvents_WorkbookBeforeCloseEventHandler(excelApp_WorkbookBeforeClose);
                
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Exception", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                if (excelApp != null)
                {
                    // excelApp.Quit();
                }
            }
            return success;
        }

        void excelApp_WorkbookBeforeSave(Excel.Workbook Wb, bool SaveAsUI, ref bool Cancel)
        {
            Wb.Application.FileDialog.InitialFileName = "ABC.xls";
        }

        void employeeWorkBook_BeforeSave(bool SaveAsUI, ref bool Cancel)
        {
            
        }

        void excelApp_WorkbookBeforeClose(Excel.Workbook wb, ref bool Cancel)
        {
           
            
            wb.Application.Quit();
            
        }

        void Application_WorkbookBeforeSave(Excel.Workbook Wb, bool SaveAsUI, ref bool Cancel)
        {
            
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


        #endregion
    }
}
