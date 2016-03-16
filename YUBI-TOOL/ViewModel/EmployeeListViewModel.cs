using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Windows;
using Caliburn.Micro;
using Microsoft.Office.Interop.Excel;
using YUBI_TOOL.Common;
using YUBI_TOOL.Model;
using YUBI_TOOL.Service;
namespace YUBI_TOOL.ViewModel
{
    public class EmployeeListViewModel : ViewModelBase
    {
        private const string FORM_ID = "EmployeeList";
        private const string COUNT_PRESENT = "=COUNTIF({2}{0}:{2}{1},\"o\") + SUM(COUNTIF({2}{0}:{2}{1},\"1/2{3}\")/2)";
        private const string COUNT_ABSENT = "=SUM(COUNTIF({11}{0}:{11}{1},{2}\"{10}\",\"{4}\",\"{5}\",\"{6}\",\"{7}\",\"{8}\"{3})) + SUM(COUNTIF({11}{0}:{11}{1},\"1/2{9}\")/2)";
        private const string COUNT_ABSENT_DETAIL = "=COUNTIF({3}{0}:{3}{1},\"{2}\")";
        private const string COUNT_ABSENT_DETAIL_P = "=SUM(COUNTIF({3}{0}:{3}{1},\"{2}\") + COUNTIF({3}{0}:{3}{1},\"1/2{2}\")/2)";
        private const string RECAPITOLATION_1_SPINNING = "RECAP SM-1";
        private const string RECAPITOLATION_1_SKINTTING = "RECAP KM";
        private const string RECAPITOLATION_2_SPINNING = "RECAP SM-2";
        private const string RECAPITOLATION_2_SKINTTING = "RECAP KNITTING";
        private const string RECAPITOLATION_EXPEDITION = "RECAP EXPEDITION";
        private const string RECAPITOLATION_CONFECTION = "RECAP CONFECTION";

        private readonly ICompanyService companyService;
        private readonly IDepartmentService departmentService;
        private readonly IWorkDataService workDataService;
        private readonly IEmployeeService employeeService;
        private readonly ITimeTableService timeTableService;
        private readonly IHolidayService holidayService;
        private readonly IMonthlyService monthlyService;
        private int colYearListIndex = 2;
        private int colMonthListIndex = 4;
        private int colDayListIndex = 6;

        private List<CompanyModel> companyList;
        private List<PostModel> departmentList;
        private CompanyModel selectedCompany;
        private PostModel selectedDepartment;
        private List<WorkDataModel> workDataList;
        private List<WorkDataModel> workDataAllList;
        private List<SelectItemModel> dayTypeList;
        private List<TimeTableModel> timeTableList;
        private List<HolidayModel> holidayList;
        private List<EmployeeModel> employeeList;
        private LanguageModel lblEmployeeList;
        private LanguageModel lblCompany;
        private LanguageModel lblDepartment;
        private LanguageModel lblEmployeeSearch;
        private LanguageModel lblMessageArea;
        private LanguageModel lblMessageText;
        private LanguageModel lblYearMonth;
        private LanguageModel lblSearch;
        private LanguageModel lblCheckAll;

        private LanguageModel lblPersonalExport;
        private LanguageModel lblDailyExport;
        private LanguageModel lblMonthlyExport;
        private LanguageModel lblCancel;
        private LanguageModel lblLogout;
        private LanguageModel lblClose;

        private string[] headerText;

        private string employeeSearch;
        private List<string> day1List;
        private List<string> day2List;
        private List<string> day3List;
        private string selectedDay1;
        private string selectedDay2;
        private string selectedDay3;
        private int selectedDay1Index;
        private int selectedDay2Index;
        private int selectedDay3Index;
        private bool canSearch;
        private bool isAllChecked;
        private bool isFocused;
        private decimal currentYearMonth;

        private WorkDataModel selectedWorkData;

        private Microsoft.Office.Interop.Excel.Application excelApp = null;
        private Workbook workBook;
        private Workbook templateWorkbook;
        private IEventAggregator eventAggregator;
        int unit_minutes = 0;
        public EmployeeListViewModel(IEventAggregator eventAggregator)
        {
            companyService = IoC.Get<ICompanyService>();
            departmentService = IoC.Get<IDepartmentService>();
            workDataService = IoC.Get<IWorkDataService>();
            employeeService = IoC.Get<IEmployeeService>();
            timeTableService = IoC.Get<ITimeTableService>();
            holidayService = IoC.Get<IHolidayService>();
            monthlyService = IoC.Get<IMonthlyService>();
            this.eventAggregator = eventAggregator;
            //this.PropertyChanged += new System.ComponentModel.PropertyChangedEventHandler(EmployeeListViewModel_PropertyChanged);
        }

        public void ReInitByDepartmentChanged()
        {
            if (canSearch)
            {
                Search();
            }
        }
        public void ReInitByCompanyChanged()
        {
            canSearch = false;
            departmentList = departmentService.SearchDepartment(CommonUtil.ToDecimal(selectedCompany.Company_no));
            selectedDepartment = departmentList.FirstOrDefault();
            var monthlyModel = monthlyService.GetMonthly(CommonUtil.ToDecimal(selectedCompany.Company_no));
            if (monthlyModel != null)
            {
                unit_minutes = (int)monthlyModel.Unit_minutes;
            }
            else
            {
                unit_minutes = 1;
            }

            Search();
            NotifyOfPropertyChange(() => DepartmentList);
            SelectedDepartment = departmentList.FirstOrDefault();
            canSearch = true;


        }
        public void ReInitByYearMonthChanged()
        {
            if (!string.IsNullOrEmpty(selectedDay1))
            {
                int selectedDay = selectedDay3Index;
                Day3List = CreateDayList(SelectedDay1, GetDayMonth(selectedDay2Index));
                if (selectedDay > day3List.Count)
                {
                    selectedDay = day3List.Count - 1;
                }
                SelectedDay3Index = selectedDay;
                if (canSearch)
                {
                    Search();
                }
            }
        }
        public void ReInitByMonthChanged()
        {
        }
        private void EmployeeListViewModel_PropertyChanged(object sender, System.ComponentModel.PropertyChangedEventArgs e)
        {
            if (e.PropertyName == GetPropertyName<EmployeeListViewModel, string>(x => x.SelectedDay1)
                || e.PropertyName == GetPropertyName<EmployeeListViewModel, int>(x => x.SelectedDay2Index))
            {
                if (!string.IsNullOrEmpty(selectedDay1))
                {
                    int selectedDay = selectedDay3Index;
                    Day3List = CreateDayList(SelectedDay1, GetDayMonth(selectedDay2Index));
                    if (selectedDay > day3List.Count)
                    {
                        selectedDay = day3List.Count - 1;
                    }
                    SelectedDay3Index = selectedDay;
                    if (canSearch)
                    {
                        Search();
                    }
                }
            }
            else if (e.PropertyName == GetPropertyName(() => SelectedCompany))
            {
                InitDepartment();
                var monthlyModel = monthlyService.GetMonthly(CommonUtil.ToDecimal(selectedCompany.Company_no));
                if (monthlyModel != null)
                {
                    unit_minutes = (int)monthlyModel.Unit_minutes;
                }
                else
                {
                    unit_minutes = 1;
                }
            }
            else if (e.PropertyName == GetPropertyName(() => SelectedDepartment) && SelectedDepartment != null)
            {
                if (canSearch)
                {
                    Search();
                }
            }


        }

        public void DoCheckAll()
        {
            if (IsAllChecked)
            {
                this.WorkDataList.ForEach(x => x.IsSelected = true);
            }
            else
            {
                this.WorkDataList.ForEach(x => x.IsSelected = false);
            }
        }
        public void DoCheck()
        {
            if (WorkDataList.Count(x => x.IsSelected) == WorkDataList.Count)
            {
                IsAllChecked = true;
            }
            else
            {
                IsAllChecked = false;
            }
        }

        public void ActiveAttendanceList()
        {
            string year;
            string month;
            year = SelectedDay1;
            month = CommonUtil.ToString(selectedDay2Index + 1);

            AttendanceListViewModel attendanceListViewModel = new AttendanceListViewModel();
            attendanceListViewModel.Init(selectedWorkData.Employee_no, selectedWorkData.Company_no, selectedWorkData.Post_no, unit_minutes, CommonUtil.GetDateAsDecimal(year, month, "1"));
            base.ActiveScreen(this, attendanceListViewModel);

        }
        protected override void OnActivate()
        {
            base.OnActivate();
            IsFocused = true;
            if (!IsActivated)
            {
                IsActivated = true;
                canSearch = false;
                InitCompany();
                InitDepartment();
                InitDaysList();
                canSearch = true;
                timeTableList = timeTableService.SearchTimeTableList();
                dayTypeList = CommonUtil.GetConstWorkDayTypeList();

            }
            Search();
        }
        protected override void OnDeactivate(bool close)
        {
            base.OnDeactivate(close);
            if (close)
            {
                Dispose();
            }
        }

        public override void CanClose(Action<bool> callback)
        {
            base.CanClose(callback);
        }
        private void InitCompany()
        {
            var companies = companyService.SearchCompanyList();
            companies.RemoveAll(x => x.Company_no == DBConstant.COMPANY_NO_ALL);
            CompanyList = companies;
            SelectedCompany = CompanyList.FirstOrDefault();
            var monthlyModel = monthlyService.GetMonthly(CommonUtil.ToDecimal(selectedCompany.Company_no));
            if (monthlyModel != null)
            {
                unit_minutes = (int)monthlyModel.Unit_minutes;
            }
            else
            {
                unit_minutes = 1;
            }
        }

        private void InitDepartment()
        {
            if (SelectedCompany != null)
            {
                DepartmentList = departmentService.SearchDepartment(CommonUtil.ToDecimal(selectedCompany.Company_no));
                SelectedDepartment = DepartmentList.FirstOrDefault();
            }
        }
        private void InitDaysList()
        {
            List<string> yearList = new List<string>();
            List<string> monthList = new List<string>();
            List<string> dayList = new List<string>();
            string currentYear = DateTime.Now.Year.ToString();
            var constform = ResourcesManager.GetLanguageForForm("const");
            var monthLong = ResourcesManager.GetLanguageForControlInForm(constform, "month_long");
            monthList = monthLong.Text.Split(',').ToList();
            for (int i = DBConstant.PROGRAM_START_YEAR; i <= DateTime.Now.Year; i++)
            {
                yearList.Add(i.ToString());
            }

            Day1List = yearList;
            Day2List = monthList;
            SelectedDay2Index = DateTime.Now.Month - 1;
            SelectedDay2 = Day2List[SelectedDay2Index];
            SelectedDay1 = currentYear;
            Day3List = CreateDayList(selectedDay1, GetDayMonth(SelectedDay2Index));
            SelectedDay3Index = DateTime.Now.Day - 1;
        }

        private string GetDayMonth(int index)
        {
            return (index + 1).ToString("00");
        }

        private static List<string> correctedDayList = new List<string>();
        private bool AddCorrectedDay(decimal date)
        {
            string dayByCompany = selectedCompany.Company_no + date;
            if (!correctedDayList.Contains(dayByCompany))
            {
                correctedDayList.Add(dayByCompany);
                return true;
            }
            return false;
        }
        public void ReInitByDateChanged()
        {
            // DO nothing
        }
        public void CorrectThirdShift()
        {
            decimal sysDate = CommonUtil.GetCurrentDate();
            string year = SelectedDay1;
            int month = selectedDay2Index + 1;
            DateTime dateFrom = new DateTime(int.Parse(year), month, 1);
            DateTime dateTo = dateFrom.AddMonths(1).AddDays(-1);
            decimal firstDayOfMonth = CommonUtil.ToDecimal(dateFrom);
            decimal lastDayOfMonth = CommonUtil.ToDecimal(dateTo);
            currentYearMonth = CommonUtil.GetDateAsDecimal(selectedDay1, CommonUtil.ToString(selectedDay2Index + 1), CommonUtil.ToString(selectedDay3Index + 1));
            var workDataToCorrectList = workDataService.SearchWorkDataListNightShift(CommonUtil.ToDecimal(selectedCompany.Company_no), firstDayOfMonth, lastDayOfMonth);
            if (workDataToCorrectList.Count > 0)
            {
                try
                {
                    string templateFile = Properties.Settings.Default.XLS_Correct_Night_Shift;
                    if (Properties.Settings.Default.XLS_Use_Multi_Language)
                    {
                        templateFile = string.Format(templateFile, Properties.Settings.Default.SelectedLanguage);
                    }
                    else
                    {
                        templateFile = string.Format(templateFile, string.Empty);
                    }

                    templateFile = CommonUtil.GetTemplate(templateFile);
                    string outFileName = string.Format("Corrected Night Shift, {0}, {1}.xls", selectedCompany.Company_name, CommonUtil.GetDateAsStringWithFormat(currentYearMonth, "dd.MM.yyyy") + ".");
                    CreateExcelApplication();
                    excelApp.Visible = false;
                    workBook = CommonUtil.CreateWorkbook(excelApp, outFileName);
                    templateWorkbook = excelApp.Workbooks.Open(templateFile, 0, false, 5, "", "", false, XlPlatform.xlWindows, "", true, false, 0, true, false, false);
                    Worksheet tempSheet = templateWorkbook.Worksheets["Before Corrected"];
                    tempSheet.Copy(workBook.Worksheets[Properties.Settings.Default.XLS_Default_Sheet_Name], Type.Missing);
                    CommonUtil.DeleteDefaultSheet(workBook);
                    int startRowIndex = 6;
                    int rowIndex = startRowIndex;
                    Worksheet correctedSheet = workBook.Worksheets["Before Corrected"];
                    Range header = correctedSheet.Range["A2", Type.Missing];
                    header.Value = selectedCompany.Company_name.ToUpper();
                    Range header1 = correctedSheet.Range["A3", Type.Missing];
                    header1.Value = FormatHeaderByDay(header1.Columns[1].Value);
                    CommonUtil.CreateEmptyRow(correctedSheet, startRowIndex, workDataToCorrectList.Count * 2);
                    CommonUtil.SetXlsPageSetup(correctedSheet, XlPageOrientation.xlLandscape);
                    List<WorkDataModel> correctList = new List<WorkDataModel>();
                    int count = 0;
                    int correctCount = 0;
                    #region start loop
                    foreach (var workDataYesterday in workDataToCorrectList)
                    {
                        DateTime yesterday = CommonUtil.ToDateTime(workDataYesterday.Work_date);
                        DateTime today = yesterday.AddDays(1);
                        decimal updateDate = CommonUtil.ToDecimal(yesterday);
                        var workDataCurrentday = workDataService.SearchWorkData(workDataYesterday.Company_no, workDataYesterday.Employee_no, CommonUtil.ToDecimal(today));
                        var employee = employeeService.GetEmployee(workDataYesterday.Company_no, 0, workDataYesterday.Employee_no, updateDate, updateDate);
                        if (employee == null)
                        {
                            continue;
                        }
                        if (workDataCurrentday != null
                                && !string.IsNullOrEmpty(workDataYesterday.Update_start_time) && string.IsNullOrEmpty(workDataYesterday.Update_end_time)
                                  && string.IsNullOrEmpty(workDataCurrentday.Start_time) && !string.IsNullOrEmpty(workDataCurrentday.End_time))
                        {
                            int columnIndex = 1;
                            correctCount++;
                            var timeTableCurrent = timeTableList.Find(x => x.Time_table_no == workDataCurrentday.Time_table_no);
                            var timeTableYesterday = timeTableList.Find(x => x.Time_table_no == workDataYesterday.Time_table_no);

                            Range detailYesterdayRange = correctedSheet.get_Range(string.Format("A{0}", rowIndex++), Type.Missing);
                            detailYesterdayRange.Columns[columnIndex++] = ++count;
                            detailYesterdayRange.Columns[columnIndex++] = workDataYesterday.Work_date;
                            detailYesterdayRange.Columns[columnIndex++] = employee.Post_name;
                            detailYesterdayRange.Columns[columnIndex++] = CommonUtil.GetFullName(employee.Emsize_first_name, employee.Emsize_last_name);
                            detailYesterdayRange.Columns[columnIndex++] = workDataYesterday.Employee_no;
                            detailYesterdayRange.Columns[columnIndex++] = CommonUtil.GetPosition(employee.Remarks);
                            detailYesterdayRange.Columns[columnIndex++] = CommonUtil.GetTimeShiftDsp(timeTableYesterday);
                            detailYesterdayRange.Columns[columnIndex++] = CommonUtil.ToDispHour(workDataYesterday.Start_time);
                            detailYesterdayRange.Columns[columnIndex++] = CommonUtil.ToDispHour(workDataYesterday.End_time);

                            Range detailTodayRange = correctedSheet.get_Range(string.Format("A{0}", rowIndex++), Type.Missing);

                            columnIndex = 1;
                            detailTodayRange.Columns[columnIndex++] = ++count;
                            detailTodayRange.Columns[columnIndex++] = workDataCurrentday.Work_date;
                            detailTodayRange.Columns[columnIndex++] = employee.Post_name;
                            detailTodayRange.Columns[columnIndex++] = CommonUtil.GetFullName(employee.Emsize_first_name, employee.Emsize_last_name);
                            detailTodayRange.Columns[columnIndex++] = workDataCurrentday.Employee_no;
                            detailTodayRange.Columns[columnIndex++] = CommonUtil.GetPosition(employee.Remarks);
                            detailTodayRange.Columns[columnIndex++] = CommonUtil.GetTimeShiftDsp(timeTableCurrent);
                            detailTodayRange.Columns[columnIndex++] = CommonUtil.ToDispHour(workDataCurrentday.Start_time);
                            detailTodayRange.Columns[columnIndex] = CommonUtil.ToDispHour(workDataCurrentday.End_time);
                            if (correctCount % 2 == 0)
                            {
                                CommonUtil.SetBackColor(detailYesterdayRange, 1, columnIndex, CommonUtil.GetExcelColor("Azure"));
                                CommonUtil.SetBackColor(detailTodayRange, 1, columnIndex, CommonUtil.GetExcelColor("Azure"));
                            }

                            TimeTableModel yesterdayShift = timeTableList.Find(x => x.Time_table_no == workDataYesterday.Time_table_no);
                            if (yesterdayShift != null && yesterdayShift.Work_from <= 2400M && (yesterdayShift.Work_to > 2400 || yesterdayShift.Work_to < yesterdayShift.Work_from))
                            {
                                decimal end_time = CommonUtil.ToDecimal(workDataCurrentday.End_time);
                                end_time = CommonUtil.MinuteToHr(CommonUtil.AddTime(end_time, 2400M));
                                workDataYesterday.End_time = end_time.ToString();
                                workDataYesterday.Update_end_time = end_time.ToString();
                                workDataYesterday.Memo = workDataYesterday.Memo + workDataCurrentday.Memo;
                                workDataCurrentday.Memo = null;
                                if (workDataYesterday.Work_type_no == DBConstant.WORK_TYPE_NORMAL
                                    || workDataYesterday.Work_type_no == DBConstant.WORK_TYPE_HOLIDAY_DUTY || workDataYesterday.Work_type_no == DBConstant.WORK_TYPE_HALF_PERMISSION)
                                {

                                    decimal? update_start_time = CommonUtil.ToNullableDecimal(workDataYesterday.Update_start_time);
                                    decimal? update_end_time = CommonUtil.ToNullableDecimal(workDataYesterday.Update_end_time);
                                    if (update_start_time != null && update_end_time != null)
                                    {

                                        string restFrom = "Rest{0}_from";
                                        string restTo = "Rest{0}_to";
                                        decimal restTime = 0;
                                        decimal overTime = 0;
                                        decimal lateness = 0;
                                        decimal leaveEarly = 0;

                                        if (yesterdayShift.Work_from < update_start_time)
                                        {
                                            lateness = CommonUtil.SubTime(update_start_time.Value, yesterdayShift.Work_from);
                                        }
                                        if (yesterdayShift.Work_to > update_end_time)
                                        {
                                            leaveEarly = CommonUtil.SubTime(yesterdayShift.Work_to, update_end_time.Value);
                                        }
                                        for (int i = 1; i <= 10; i++)
                                        {
                                            decimal? restTimeFrom = (decimal?)CommonUtil.GetPropertyValue(yesterdayShift, string.Format(restFrom, i));
                                            decimal? restTimeTo = (decimal?)CommonUtil.GetPropertyValue(yesterdayShift, string.Format(restTo, i));

                                            if (restTimeFrom != null && restTimeTo != null)
                                            {
                                                if (restTimeFrom < yesterdayShift.Work_from && restTimeTo < yesterdayShift.Work_from)
                                                {
                                                    restTimeFrom = CommonUtil.MinuteToHr(CommonUtil.AddTime(restTimeFrom.Value, 2400M));
                                                    restTimeTo = CommonUtil.MinuteToHr(CommonUtil.AddTime(restTimeTo.Value, 2400M));
                                                }
                                                decimal over_Time_at_start = 0;
                                                decimal over_Time_at_end = 0;

                                                // start work
                                                if (restTimeFrom.Value < yesterdayShift.Work_from)
                                                {
                                                    if (update_start_time < restTimeFrom.Value)
                                                    {
                                                        restTime += CommonUtil.SubTime(restTimeTo.Value, restTimeFrom.Value);
                                                        over_Time_at_start = CommonUtil.SubTime(restTimeFrom.Value, update_start_time.Value);
                                                    }
                                                    else if (lateness == 0)
                                                    {
                                                        restTime += CommonUtil.SubTime(yesterdayShift.Work_from, update_start_time.Value);
                                                    }
                                                }
                                                // end work
                                                else if (restTimeTo > yesterdayShift.Work_to)
                                                {
                                                    if (update_end_time > restTimeTo.Value)
                                                    {
                                                        restTime += CommonUtil.SubTime(restTimeTo.Value, restTimeFrom.Value);
                                                        over_Time_at_end = CommonUtil.SubTime(update_end_time.Value, restTimeTo.Value);
                                                    }
                                                    else if (leaveEarly == 0)
                                                    {
                                                        // not late
                                                        restTime += CommonUtil.SubTime(update_end_time.Value, yesterdayShift.Work_to);
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
                                            workDataYesterday.Leaving_early_time = CommonUtil.ToString(leaveEarly);
                                        }
                                        else
                                        {
                                            workDataYesterday.Leaving_early_time = null;
                                        }
                                        if (lateness > 0)
                                        {
                                            workDataYesterday.Being_late_time = CommonUtil.ToString(lateness);
                                        }
                                        else
                                        {
                                            workDataYesterday.Being_late_time = null;
                                        }
                                        workDataYesterday.Rest_time = CommonUtil.ToString(restTime);
                                        workDataYesterday.Over_time = CommonUtil.ToString(overTime);
                                        decimal working_time = CommonUtil.SubTime(yesterdayShift.Work_to, yesterdayShift.Work_from) - leaveEarly - lateness + overTime;
                                        if (workDataYesterday.Work_type_no == DBConstant.WORK_TYPE_HOLIDAY_DUTY)
                                        {
                                            workDataYesterday.Working_time = null;
                                            workDataYesterday.Holiday_time = CommonUtil.ToString(working_time);
                                            workDataYesterday.Holiday_days = 1M;
                                        }
                                        else
                                        {
                                            workDataYesterday.Holiday_time = null;
                                            workDataYesterday.Working_time = CommonUtil.ToString(working_time);
                                            workDataYesterday.Holiday_days = null;
                                        }

                                        if (workDataYesterday.Work_type_no == DBConstant.WORK_TYPE_NORMAL && !workDataYesterday.IsHoliday)
                                        {
                                            workDataYesterday.Contract_time = CommonUtil.GetContactTime(yesterdayShift).ToString();
                                            workDataYesterday.Work_days = 1M;
                                        }
                                        else if (workDataYesterday.Work_type_no == DBConstant.WORK_TYPE_HALF_PERMISSION)
                                        {
                                            workDataYesterday.Contract_time = CommonUtil.DivideTime(CommonUtil.GetContactTime(yesterdayShift), 2).ToString();
                                            workDataYesterday.Late_night_time = null;
                                            workDataYesterday.Leaving_early_days = null;
                                            workDataYesterday.Leaving_early_time = null;
                                            workDataYesterday.Work_days = 0.5M;
                                        }
                                        else
                                        {
                                            workDataYesterday.Contract_time = "0";
                                            workDataYesterday.Holiday_days = null;
                                            workDataYesterday.Work_days = null;
                                        }
                                    }
                                }

                                workDataCurrentday.End_time = null;
                                workDataCurrentday.Update_end_time = null;
                                correctList.Add(workDataYesterday);
                                correctList.Add(workDataCurrentday);
                            }

                        }

                    }
                    #endregion loop
                    templateWorkbook.Close();
                    Show();
                    if (correctList.Count > 0)
                    {
                        workDataService.UpdateWorkDataList(correctList);

                    }

                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                    excelApp.Quit();
                }
                finally
                {
                    DisposeXls();
                }
            }

        }

        private void CorrectUnvalidWorkdate()
        {
            if (workDataAllList != null)
            {
                decimal sysDate = CommonUtil.GetCurrentDate();
                decimal currentTime = decimal.Parse(string.Format("{0}{1}", DateTime.Now.TimeOfDay.Hours.ToString("00"), DateTime.Now.TimeOfDay.Minutes.ToString("00")));
                var invalidWorkList = workDataAllList.FindAll(x => (x.IsHoliday
                    && x.Work_type_no == DBConstant.WORK_TYPE_NORMAL
                    && !string.IsNullOrEmpty(x.Update_start_time)
                    && string.IsNullOrEmpty(x.Update_end_time))
                    || (x.Work_from >= 1800 && !string.IsNullOrEmpty(x.Update_start_time))
                    && CommonUtil.ToDecimal(x.Update_start_time) <= 1200 && string.IsNullOrEmpty(x.Update_end_time));
                invalidWorkList.ForEach(delegate(WorkDataModel work)
                {
                    work.Memo = string.Format("Removed unvalid working Time: AT[{0}]", CommonUtil.ToDispHour(work.Update_start_time));
                    work.Start_time = null;
                    work.End_time = null;
                    work.Update_start_time = null;
                    work.Update_end_time = null;
                    work.Working_time = null;
                    work.Over_time = null;
                    work.Holiday_time = null;
                    if (!work.IsHoliday && work.WorkingType == WorkingType.Present
                        && work.Work_date == sysDate && currentTime < work.Work_from - 15)
                    {
                        work.WorkingType = WorkingType.Unknown;
                    }
                });
                if (invalidWorkList.Count > 0)
                {
                    workDataService.UpdateWorkDataList(invalidWorkList);
                }
            }
        }
        public void Search()
        {
            if (this.workDataList != null)
            {
                CommonUtil.ClearSortDirection(this.workDataList);
            }
            decimal systemDate = CommonUtil.GetCurrentDate();
            string year = SelectedDay1;
            int month = selectedDay2Index + 1;
            DateTime dateFrom = new DateTime(int.Parse(year), month, 1);
            DateTime dateTo = dateFrom.AddMonths(1).AddDays(-1);
            decimal firstDayOfMonth = CommonUtil.ToDecimal(dateFrom);
            decimal lastDayOfMonth = CommonUtil.ToDecimal(dateTo);
            holidayList = holidayService.SearchHolidayList(CommonUtil.ToDecimal(selectedCompany.Company_no), firstDayOfMonth, lastDayOfMonth);
            employeeList = employeeService.SearchEmployeeList(CommonUtil.ToDecimal(SelectedCompany.Company_no), CommonUtil.ToDecimal(this.selectedDepartment.Post_no), employeeSearch, firstDayOfMonth, lastDayOfMonth);
            // create data in month
            if (Properties.Settings.Default.AutoCorrectThirdShift)
            {
                CorrectThirdShift();
            }
            workDataAllList = workDataService.SearchWorkDataList(CommonUtil.ToDecimal(SelectedCompany.Company_no), CommonUtil.ToDecimal(this.selectedDepartment.Post_no), employeeSearch, firstDayOfMonth, lastDayOfMonth);
            if (workDataAllList.Count == 0)
            {
                foreach (var employee in employeeList)
                {
                    workDataService.CreateWorkDataInMonth(employee, firstDayOfMonth);
                }
                workDataAllList = workDataService.SearchWorkDataList(CommonUtil.ToDecimal(SelectedCompany.Company_no), CommonUtil.ToDecimal(this.selectedDepartment.Post_no), employeeSearch, firstDayOfMonth, lastDayOfMonth);
            }
            List<WorkDataModel> workDataList = new List<WorkDataModel>();
            foreach (var employee in employeeList)
            {
                string position = CommonUtil.GetPosition(employee.Remarks);
                var workDatas = workDataAllList.FindAll(x => x.Employee_no == employee.Employee_no && x.Company_no == employee.Company_no && x.Post_no == employee.Post_no);
                decimal contactHr = 0;
                decimal workingHr = 0;
                decimal overTimeHr = 0;
                string memo = "";
                workDatas.ForEach(delegate(WorkDataModel work)
                {
                    bool isOutOfExpiration = false;
                    if (employee.Expiration_to > 0 && employee.Expiration_to < 99999999M)
                    {
                        DateTime expirationTo = CommonUtil.ToDateTime(employee.Expiration_to);

                        if (expirationTo.Day == 1 && employee.Expiration_to <= firstDayOfMonth)
                        {
                            isOutOfExpiration = true;
                        }

                    }
                    work.WorkingDayType = CommonUtil.GetWorkingDayType(work, employee, holidayList);

                    if (work.WorkingDayType == WorkingDayType.NationalHoliday || work.WorkingDayType == WorkingDayType.NormalHoliday)
                    {
                        work.IsHoliday = true;
                    }
                    work.WorkingType = CommonUtil.GetWorkingType(work);
                    SetTimeTable(work);

                    if (!isOutOfExpiration && employee.Expiration_from <= work.Work_date && employee.Expiration_to >= work.Work_date)
                    {

                        work.IsOutOfExpiration = false;

                    }
                    else
                    {
                        work.IsOutOfExpiration = true;
                    }

                    if (work.Contract_time != null && !work.IsOutOfExpiration && !work.IsHoliday && (work.Work_type_no == DBConstant.WORK_TYPE_NORMAL
                                                || work.Work_type_no == DBConstant.WORK_TYPE_HALF_PERMISSION))
                    {
                        contactHr += CommonUtil.ToDecimal(work.Contract_time);
                    }

                    if (work.Working_time != null)
                    {
                        workingHr += CommonUtil.ToDecimal(work.Working_time);
                    }
                    if (work.Over_time != null)
                    {
                        overTimeHr += CommonUtil.ToDecimal(work.Over_time);
                    }
                    work.IsHasNoOnOffDuty = false;
                    if (!work.IsOutOfExpiration &&
                        ((work.Work_type_no == DBConstant.WORK_TYPE_NORMAL
                        || work.Work_type_no == DBConstant.WORK_TYPE_HALF_PERMISSION) && !work.IsHoliday
                        || (work.Work_type_no == DBConstant.WORK_TYPE_HOLIDAY_DUTY) && work.IsHoliday)
                        && work.Work_date <= systemDate)
                    {
                        // already finish the work date
                        if (work.Work_date < systemDate)
                        {
                            if (string.IsNullOrEmpty(work.Update_start_time) || string.IsNullOrEmpty(work.Update_end_time))
                            {
                                work.IsHasNoOnOffDuty = true;
                            }
                            if (string.IsNullOrEmpty(work.Start_time))
                            {
                                work.IsNoOnDuty = true;
                            }
                            if (string.IsNullOrEmpty(work.End_time))
                            {
                                work.IsNoOffDuty = true;
                            }
                        }
                        else
                        {
                            // current work date or future
                            bool isBefore = false;

                            bool isInWorkingTime = IsInOrBeforeWorkingTime(work, out isBefore);

                            if (isInWorkingTime)
                            {
                                // in working
                                if (string.IsNullOrEmpty(work.Start_time))
                                {
                                    work.IsNoOnDuty = true;
                                }
                                if (string.IsNullOrEmpty(work.Update_start_time))
                                {
                                    work.IsHasNoOnOffDuty = true;
                                    if (string.IsNullOrEmpty(work.Update_end_time) && work.WorkingType == WorkingType.Present)
                                    {
                                        work.WorkingType = WorkingType.Unknown;
                                    }
                                }

                            }
                            else if (!isBefore)
                            {
                                // finish working
                                if (string.IsNullOrEmpty(work.End_time))
                                {
                                    work.IsNoOffDuty = true;

                                }

                                if (string.IsNullOrEmpty(work.Start_time))
                                {
                                    work.IsNoOnDuty = true;

                                }
                                if (string.IsNullOrEmpty(work.Update_start_time) || string.IsNullOrEmpty(work.Update_end_time))
                                {
                                    work.IsHasNoOnOffDuty = true;
                                }
                            }
                            else if (work.WorkingType == WorkingType.Present && string.IsNullOrEmpty(work.Update_start_time))
                            {
                                // unstart working
                                work.WorkingType = WorkingType.Unknown;
                            }
                        }

                    }
                    if (work.WorkingType == WorkingType.Present && string.IsNullOrEmpty(work.Update_start_time)
                        && string.IsNullOrEmpty(work.Update_end_time))
                    {
                        work.WorkingType = WorkingType.Unknown;
                    }
                    if (work.IsHasNoOnOffDuty && string.IsNullOrEmpty(memo))
                    {
                        var messageModel = ResourcesManager.GetMessage(MessageConstant.A0062);
                        memo = messageModel.Message;
                    }
                    work.Position = position;
                });
                decimal contactHrEven = decimal.Floor(contactHr / 60);
                decimal contactHrOdd = contactHr - contactHrEven * 60;
                contactHr = contactHrEven + contactHrOdd / 100;
                decimal workingHrEven = decimal.Floor(workingHr / 60);
                decimal workingHrOdd = workingHr - workingHrEven * 60;
                workingHr = workingHrEven + workingHrOdd / 100;
                decimal overTimeHrEven = decimal.Floor(overTimeHr / 60);
                decimal overTimeHrOdd = overTimeHr - overTimeHrEven * 60;
                overTimeHr = overTimeHrEven + overTimeHrOdd / 100;
                WorkDataModel workDataModel = new WorkDataModel()
                {
                    Employee_no = employee.Employee_no,
                    EmployeeName = CommonUtil.GetFullName(employee.Emsize_first_name, employee.Emsize_last_name),
                    Contract_time = CommonUtil.ToString(contactHr),
                    Working_time = CommonUtil.ToString(workingHr),
                    Over_time = CommonUtil.ToString(overTimeHr),
                    Post_name = employee.Post_name,
                    Memo = memo,
                    Company_no = employee.Company_no,
                    Post_no = employee.Post_no,
                };
                workDataList.Add(workDataModel);

            }
            CorrectUnvalidWorkdate();
            this.WorkDataList = workDataList;
            workDataList.ForEach(x => x.IsSelected = IsAllChecked);
            Message = new MessageModel();
        }
        private List<string> CreateDayList(string year, string month)
        {
            int yearInt = int.Parse(year);
            int monthInt = int.Parse(month);
            List<string> days = new List<string>();
            DateTime yearMonth = new DateTime(yearInt, monthInt, 1);
            DateTime lastDayOfMonth = yearMonth.AddMonths(1).AddDays(-1); ;
            for (int i = yearMonth.Day; i <= lastDayOfMonth.Day; i++)
            {
                days.Add(i.ToString("00"));
            }
            return days;
        }

        public void Logout()
        {
            eventAggregator.Publish(new LogoutEvent());
        }
        protected override void SetMultiLanguage()
        {
            base.SetMultiLanguage();
            LanguageModel employeeListForm = ResourcesManager.GetLanguageForForm(FORM_ID);
            var display = ResourcesManager.GetLanguageForControlInForm(employeeListForm, "Text");
            if (display != null)
            {
                SetDisplayName(display.Text);
            }

            LblEmployeeList = ResourcesManager.GetLanguageForControlInForm(employeeListForm, "Title");
            LblMessageArea = ResourcesManager.GetLanguageForControlInForm(employeeListForm, "MessageArea");
            LblMessageText = ResourcesManager.GetLanguageForControlInForm(employeeListForm, "MessageText");
            LblCheckAll = ResourcesManager.GetLanguageForControlInForm(employeeListForm, "chkExport");

            LblCompany = ResourcesManager.GetLanguageForControlInForm(employeeListForm, "Label3");
            LblDepartment = ResourcesManager.GetLanguageForControlInForm(employeeListForm, "Label1");
            LblEmployeeSearch = ResourcesManager.GetLanguageForControlInForm(employeeListForm, "Label8");
            LblSearch = ResourcesManager.GetLanguageForControlInForm(employeeListForm, "btn_Search");
            LblYearMonth = ResourcesManager.GetLanguageForControlInForm(employeeListForm, "Label2");
            var dataGridView = ResourcesManager.GetLanguageForControlInForm(employeeListForm, "DataGridView");
            var dataGridHeader = ResourcesManager.GetLanguageForControlInForm(dataGridView, "HeaderText");
            if (dataGridHeader != null && !string.IsNullOrEmpty(dataGridHeader.Text))
            {
                HeaderText = dataGridHeader.Text.Split(',');
            }
            LblCancel = ResourcesManager.GetLanguageForControlInForm(employeeListForm, "btnPrevious");

            LblDailyExport = ResourcesManager.GetLanguageForControlInForm(employeeListForm, "btnAttendanceListXlsExport");
            LblMonthlyExport = ResourcesManager.GetLanguageForControlInForm(employeeListForm, "btnCustomizeList1XlsExport");
            LblPersonalExport = ResourcesManager.GetLanguageForControlInForm(employeeListForm, "btnXlsExport");
            LblLogout = ResourcesManager.GetLanguageForForm("Menu", "btnLogout");
            LblClose = ResourcesManager.GetLanguageForForm("Login", "btnEnd");
            if (Properties.Settings.Default.SelectedLanguage != "Japanese")
            {
                ColDayListIndex = 2;
                ColMonthListIndex = 4;
                ColYearListIndex = 6;
            }
        }
        private void ParseData(List<WorkDataModel> workDataList, List<HolidayModel> holidayList, EmployeeModel employeeModel)
        {
            foreach (var workData in workDataList)
            {
                workData.Work_date_dsp = CommonUtil.GetDateAsString(workData.Work_date, ResourcesManager.LIST_DATE_FORMAT, true);
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

                SetTimeTable(workData);
            }
        }

        private string CreateFileName(List<WorkDataModel> workList)
        {
            string employeeNos = "";
            int countNo = 0;
            bool hasUnknowEmployeeNo = false;

            foreach (var workData in workList)
            {
                countNo++;
                if (countNo <= DBConstant.MAX_EMPLOYEE_NO_PER_FILE_NAME)
                {
                    employeeNos += workData.Employee_no + ",";
                }
                else if (countNo > DBConstant.MAX_EMPLOYEE_NO_PER_FILE_NAME && countNo == workList.Count)
                {
                    employeeNos += workData.Employee_no;
                }
                else if (!hasUnknowEmployeeNo)
                {
                    employeeNos += "...,";
                    hasUnknowEmployeeNo = true;
                }
            }

            return employeeNos.TrimEnd(',');
        }

        public void PersonalExport()
        {
            Message = null;
            currentYearMonth = CommonUtil.GetDateAsDecimal(selectedDay1, CommonUtil.ToString(selectedDay2Index + 1), CommonUtil.ToString(selectedDay3Index + 1));

            var checkedList = CommonUtil.GetSortedList(workDataList).FindAll(x => x.IsSelected);
            if (checkedList.Count == 0)
            {
                Message = ResourcesManager.GetMessage(MessageConstant.A0011);
                return;
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
            decimal firstDayOfMonth = CommonUtil.GetFirstDayOfMonth(currentYearMonth);
            decimal lastDayOfMonth = CommonUtil.GetLastDayOfMonth(currentYearMonth);

            templateFile = CommonUtil.GetTemplate(templateFile);
            bool hasData = true;

            do
            {
                string outFileName = "";
                List<WorkDataModel> wkList;
                if (checkedList.Count > DBConstant.MAX_SHEET_PER_WORKSPACE)
                {
                    wkList = checkedList.GetRange(0, DBConstant.MAX_SHEET_PER_WORKSPACE);
                    checkedList.RemoveRange(0, DBConstant.MAX_SHEET_PER_WORKSPACE);
                }
                else
                {
                    wkList = checkedList;
                    hasData = false;
                }

                try
                {
                    outFileName = CreateFileName(wkList);
                    outFileName = string.Format(Properties.Settings.Default.XLS_Out_Personal_File, outFileName, CommonUtil.GetDateAsString(currentYearMonth, "MY").ToUpper() + ".");
                    CreateExcelApplication();
                    templateWorkbook = excelApp.Workbooks.Open(templateFile, 0, false, 5, "", "", false, XlPlatform.xlWindows, "", true, false, 0, true, false, false);
                    Worksheet tempSheet = templateWorkbook.Worksheets["Employee_no"];
                    workBook = CommonUtil.CreateWorkbook(excelApp, outFileName);
                    foreach (var workData in wkList)
                    {

                        List<WorkDataModel> exportList = workDataService.SearchWorkDataListByEmployee(workData.Company_no, 0, workData.Employee_no, firstDayOfMonth, lastDayOfMonth);
                        EmployeeModel employeeModel = employeeList.Find(x => x.Employee_no == workData.Employee_no);
                        tempSheet.Copy(workBook.Worksheets[Properties.Settings.Default.XLS_Default_Sheet_Name], Type.Missing);
                        Worksheet employeeSheet = (Worksheet)workBook.Worksheets["Employee_no"];
                        CommonUtil.SetXlsPageSetup(employeeSheet, XlPageOrientation.xlLandscape);
                        employeeSheet.Name = employeeModel.Employee_no;
                        ParseData(exportList, holidayList, employeeModel);
                        CommonUtil.FillPersonalReportSheet(employeeSheet, employeeModel, exportList, dayTypeList, timeTableList, unit_minutes);
                    }

                    CommonUtil.DeleteDefaultSheet(workBook);
                    templateWorkbook.Saved = true;
                    templateWorkbook.Close();
                    Show();

                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                    excelApp.Quit();
                }
                finally
                {
                    DisposeXls();
                }
            } while (hasData);




        }

        private bool IsInOrBeforeWorkingTime(WorkDataModel workData, out bool isBefore)
        {
            bool isInWorkingTime = false;
            isBefore = false;
            TimeTableModel timeTable = timeTableList.Find(x => x.Time_table_no == workData.Time_table_no);
            decimal currentTime = decimal.Parse(string.Format("{0}{1}", DateTime.Now.TimeOfDay.Hours.ToString("00"), DateTime.Now.TimeOfDay.Minutes.ToString("00")));
            decimal currentDate = CommonUtil.GetCurrentDate();
            if (timeTable != null)
            {
                if (currentTime < timeTable.Work_from)
                {
                    isBefore = true;
                }
                else
                {
                    string rest_to_pattern = "Rest{0}_to";
                    if (timeTable.Work_to > 2400 || timeTable.Work_to < timeTable.Work_from)
                    {
                        // night shifts
                        decimal workTo = CommonUtil.MinuteToHr(CommonUtil.SubTime(timeTable.Work_to, 2400));
                        decimal maxRestTo = workTo;
                        for (int i = 1; i <= 10; i++)
                        {
                            decimal? restTo = (decimal?)CommonUtil.GetPropertyValue(timeTable, string.Format(rest_to_pattern, i));
                            if (restTo != null && restTo.Value < timeTable.Work_from && restTo.Value > workTo && maxRestTo < restTo.Value)
                            {
                                maxRestTo = restTo.Value;
                            }
                        }
                        if (currentTime < maxRestTo || currentTime <= 2400)
                        {
                            isInWorkingTime = true;
                        }

                    }
                    else
                    {
                        // day shifts
                        decimal maxRestTo = timeTable.Work_to;
                        for (int i = 1; i <= 10; i++)
                        {
                            decimal? restTo = (decimal?)CommonUtil.GetPropertyValue(timeTable, string.Format(rest_to_pattern, i));
                            if (restTo != null && restTo.Value > timeTable.Work_to && maxRestTo < restTo.Value)
                            {
                                maxRestTo = restTo.Value;
                            }
                        }
                        if (currentTime < maxRestTo)
                        {
                            isInWorkingTime = true;
                        }
                    }
                }
            }
            return isInWorkingTime;
        }
        public void DailyExport()
        {
            Message = null;
            currentYearMonth = CommonUtil.GetDateAsDecimal(selectedDay1, CommonUtil.ToString(selectedDay2Index + 1), CommonUtil.ToString(selectedDay3Index + 1));
            var sortedList = CommonUtil.GetSortedList(workDataList);
            var checkedList = sortedList.FindAll(x => x.IsSelected).Select(x => x.Employee_no).ToList();
            if (checkedList.Count == 0)
            {
                Message = ResourcesManager.GetMessage(MessageConstant.A0011);
                return;
            }
            string templateFile = Properties.Settings.Default.XLS_Daily_Report;
            if (Properties.Settings.Default.XLS_Use_Multi_Language)
            {
                templateFile = string.Format(templateFile, Properties.Settings.Default.SelectedLanguage);
            }
            else
            {
                templateFile = string.Format(templateFile, string.Empty);
            }
            List<WorkDataModel> dailyList = new List<WorkDataModel>();
            foreach (var employee in checkedList)
            {
                dailyList.AddRange(workDataAllList.FindAll(x => x.Work_date == currentYearMonth && x.Employee_no == employee));
            }

            holidayList = holidayService.SearchHolidayList(CommonUtil.ToDecimal(selectedCompany.Company_no), currentYearMonth, currentYearMonth);

            templateFile = CommonUtil.GetTemplate(templateFile);

            bool isAllDepartment = false;
            string outFileName = selectedCompany.Company_name;
            if (selectedDepartment.Post_no != "0")
            {
                isAllDepartment = false;
                outFileName += " , " + selectedDepartment.Post_name;
            }
            else
            {
                isAllDepartment = true;
            }
            outFileName = CommonUtil.Capitalize(outFileName);
            outFileName = string.Format(Properties.Settings.Default.XLS_Out_Daily_File, outFileName, CommonUtil.GetDateAsStringWithFormat(currentYearMonth, "dd.MM.yyyy") + ".");

            try
            {
                CreateExcelApplication();
                excelApp.Visible = false;
                workBook = CommonUtil.CreateWorkbook(excelApp, outFileName);
                templateWorkbook = excelApp.Workbooks.Open(templateFile, 0, false, 5, "", "", false, XlPlatform.xlWindows, "", true, false, 0, true, false, false);
                Worksheet tempWorkOnduty;
                Worksheet tempWorkOnOffduty;
                Worksheet tempWorkLate;
                Worksheet tempWorkLeaveEarly;
                Worksheet tempDailySheet = null;
                Worksheet tempRecap1Sheet = null;
                Worksheet tempRecap2Sheet = null;
                Worksheet tempRecapConfectionSheet = null;
                Worksheet tempRecapExpeditionSheet = null;


                if (isAllDepartment)
                {
                    tempDailySheet = templateWorkbook.Worksheets[DBConstant.TEMP_DAILY_REPORT_BY_COMPANY];
                }
                else
                {
                    tempDailySheet = templateWorkbook.Worksheets[DBConstant.TEMP_DAILY_REPORT_BY_DEPART];
                }
                tempDailySheet.Name = DBConstant.SHEET_DAILY_REPORT;
                tempWorkOnduty = templateWorkbook.Worksheets[DBConstant.TEMP_DAILY_REPORT_NO_ON_DUTY];
                tempWorkOnOffduty = templateWorkbook.Worksheets[DBConstant.TEMP_DAILY_REPORT_NO_OFF_DUTY];
                tempWorkLate = templateWorkbook.Worksheets[DBConstant.TEMP_DAILY_REPORT_LATE];
                tempWorkLeaveEarly = templateWorkbook.Worksheets[DBConstant.TEMP_DAILY_REPORT_LEFT_EARLIER];



                tempDailySheet.Copy(workBook.Worksheets[Properties.Settings.Default.XLS_Default_Sheet_Name], Type.Missing);
                tempWorkOnduty.Copy(workBook.Worksheets[Properties.Settings.Default.XLS_Default_Sheet_Name], Type.Missing);
                tempWorkOnOffduty.Copy(workBook.Worksheets[Properties.Settings.Default.XLS_Default_Sheet_Name], Type.Missing);
                tempWorkLate.Copy(workBook.Worksheets[Properties.Settings.Default.XLS_Default_Sheet_Name], Type.Missing);
                tempWorkLeaveEarly.Copy(workBook.Worksheets[Properties.Settings.Default.XLS_Default_Sheet_Name], Type.Missing);

                Worksheet dailySheet = (Worksheet)workBook.Worksheets[tempDailySheet.Name];
                Worksheet noOndutySheet = (Worksheet)workBook.Worksheets[tempWorkOnduty.Name];
                Worksheet noOffdutySheet = (Worksheet)workBook.Worksheets[tempWorkOnOffduty.Name];
                Worksheet lateSheet = (Worksheet)workBook.Worksheets[tempWorkLate.Name];
                Worksheet leaveEarlySheet = (Worksheet)workBook.Worksheets[tempWorkLeaveEarly.Name];

                CommonUtil.SetXlsPageSetup(dailySheet, XlPageOrientation.xlPortrait);
                CommonUtil.SetXlsPageSetup(noOndutySheet, XlPageOrientation.xlPortrait);
                CommonUtil.SetXlsPageSetup(noOffdutySheet, XlPageOrientation.xlPortrait);
                CommonUtil.SetXlsPageSetup(lateSheet, XlPageOrientation.xlPortrait);
                CommonUtil.SetXlsPageSetup(leaveEarlySheet, XlPageOrientation.xlPortrait);

                SortDataByShiftForDaily(dailyList);
                SetDailyReportSheet(dailySheet, isAllDepartment, dailyList);
                dailyList = dailyList.FindAll(x => !x.IsOutOfExpiration);
                List<WorkDataModel> noOndutyList = dailyList.FindAll(x => x.IsNoOnDuty);
                SetDutySheet(tempWorkOnduty, noOndutySheet, noOndutyList);
                noOndutyList.Clear();
                List<WorkDataModel> noOffdutyList = dailyList.FindAll(x => x.IsNoOffDuty);
                SetDutySheet(tempWorkOnOffduty, noOffdutySheet, noOffdutyList);
                noOffdutyList.Clear();
                List<WorkDataModel> lateList = dailyList.FindAll(x => x.IsLate);
                SetLateEarlySheet(lateSheet, lateList);
                lateList.Clear();
                List<WorkDataModel> leavEarlyList = dailyList.FindAll(x => x.IsLeaveEarly);
                SetLateEarlySheet(leaveEarlySheet, leavEarlyList);
                leavEarlyList.Clear();

                if (isAllDepartment)
                {
                    tempRecap1Sheet = templateWorkbook.Worksheets[DBConstant.TEMP_DAILY_RECAP_1];
                    tempRecap2Sheet = templateWorkbook.Worksheets[DBConstant.TEMP_DAILY_RECAP_2];

                    tempRecap1Sheet.Copy(workBook.Worksheets[Properties.Settings.Default.XLS_Default_Sheet_Name], Type.Missing);
                    tempRecap2Sheet.Copy(workBook.Worksheets[Properties.Settings.Default.XLS_Default_Sheet_Name], Type.Missing);
                    Worksheet recap1Sheet = (Worksheet)workBook.Worksheets[tempRecap1Sheet.Name];
                    Worksheet recap2Sheet = (Worksheet)workBook.Worksheets[tempRecap2Sheet.Name];
                    CommonUtil.SetXlsPageSetup(recap1Sheet, XlPageOrientation.xlLandscape);
                    CommonUtil.SetXlsPageSetup(recap2Sheet, XlPageOrientation.xlLandscape);
                    if (selectedCompany.Company_no == DBConstant.COMPANY_NO_SPINNING_MILL)
                    {
                        recap1Sheet.Name = DBConstant.SHEET_RECAP_SPINNING_1;
                        recap2Sheet.Name = DBConstant.SHEET_RECAP_SPINNING_2;

                        SetRECAP1(tempRecap1Sheet, recap1Sheet, dailyList, false);
                        SetRECAP2(tempRecap2Sheet, recap2Sheet, dailyList, true);
                    }
                    else
                    {
                        recap1Sheet.Name = DBConstant.SHEET_RECAP_KNITTING_1;
                        recap2Sheet.Name = DBConstant.SHEET_RECAP_KNITTING_2;
                        tempRecapConfectionSheet = templateWorkbook.Worksheets[DBConstant.SHEET_RECAP_CONFECTION];
                        tempRecapExpeditionSheet = templateWorkbook.Worksheets[DBConstant.SHEET_RECAP_EXPEDITION];
                        tempRecapConfectionSheet.Copy(workBook.Worksheets[Properties.Settings.Default.XLS_Default_Sheet_Name], Type.Missing);
                        tempRecapExpeditionSheet.Copy(workBook.Worksheets[Properties.Settings.Default.XLS_Default_Sheet_Name], Type.Missing);

                        Worksheet recapConfectionSheet = (Worksheet)workBook.Worksheets[tempRecapConfectionSheet.Name];
                        Worksheet recapExpeditionSheet = (Worksheet)workBook.Worksheets[tempRecapExpeditionSheet.Name];
                        CommonUtil.SetXlsPageSetup(recapConfectionSheet, XlPageOrientation.xlLandscape);
                        CommonUtil.SetXlsPageSetup(recapExpeditionSheet, XlPageOrientation.xlLandscape);

                        SetRECAP1(tempRecap1Sheet, recap1Sheet, dailyList, true);
                        SetRECAP2(tempRecap2Sheet, recap2Sheet, dailyList.FindAll(x => x.Post_name.ToUpper().StartsWith(DBConstant.RECAP_2_START_WITH_KNITTING_DEPRT)), false);
                        List<WorkDataModel> workDataByConfection = dailyList.FindAll(x => x.Post_name.ToUpper().StartsWith(DBConstant.CONFECTION_START_WITH_KNITTING_DEPRT));
                        SetRECAPConfectionOrExpedition(tempRecapConfectionSheet, recapConfectionSheet, workDataByConfection);
                        workDataByConfection.Clear();
                        List<WorkDataModel> workDataByExpedition = dailyList.FindAll(x => x.Post_name.ToUpper().Contains(DBConstant.EXPEDITION_START_WITH_KNITTING_DEPRT));
                        SetRECAPConfectionOrExpedition(tempRecapExpeditionSheet, recapExpeditionSheet, workDataByExpedition);
                        workDataByConfection.Clear();
                    }

                }
                ((Microsoft.Office.Interop.Excel._Worksheet)dailySheet).Activate();
                dailyList.Clear();
                templateWorkbook.Saved = true;
                templateWorkbook.Close();
                CommonUtil.DeleteDefaultSheet(workBook);
                this.Show();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                excelApp.Quit();
            }
            finally
            {
                DisposeXls();
            }

        }

        private void DisposeXls()
        {

            if (templateWorkbook != null)
            {
                Marshal.ReleaseComObject(this.templateWorkbook);
            }
            if (this.workBook != null)
            {
                Marshal.ReleaseComObject(this.workBook);
            }

            if (this.excelApp != null)
            {
                Marshal.ReleaseComObject(this.excelApp);
            }
            GC.Collect();
        }

        private void SetDailyReportSheet(Worksheet dailySheet, bool isAllDepartment, List<WorkDataModel> workDataList)
        {
            Range headerCompanyRange = dailySheet.get_Range("A2", Type.Missing);
            Range headerReportNameRange = dailySheet.get_Range("A3", Type.Missing);
            headerReportNameRange.Columns[1].Value = FormatHeaderByDay(headerReportNameRange.Columns[1].Value);
            headerCompanyRange.Columns[1].Value = selectedCompany.Company_name;

            Range detailRange = dailySheet.get_Range("A6", Type.Missing);
            Range detailTemRange = dailySheet.get_Range("A6", Type.Missing);

            int detailStartRowIndex = 6;
            int rowIndex = detailStartRowIndex;
            string detailRowStart = "A{0}";
            int count = 0;
            int endRow = workDataList.Count + detailStartRowIndex;
            if (endRow >= detailStartRowIndex)
            {
                endRow--;
            }
            CommonUtil.CreateEmptyRow(dailySheet, detailStartRowIndex, workDataList.Count);
            foreach (var workData in workDataList)
            {
                int columnIndex = 1;
                Range detailEditRange = dailySheet.get_Range(string.Format(detailRowStart, rowIndex), Type.Missing).EntireRow;

                detailEditRange.Columns[columnIndex++].Value = ++count;
                if (isAllDepartment)
                {
                    detailEditRange.Columns[columnIndex++].Value = CommonUtil.GetXlsValue(workData.Post_name);
                }
                string work = CommonUtil.GetPresence(workData);

                detailEditRange.Columns[columnIndex++].Value = workData.EmployeeName;
                detailEditRange.Columns[columnIndex++].Value = workData.Employee_no;
                detailEditRange.Columns[columnIndex++].Value = CommonUtil.GetPosition(workData.Employee_remarks);
                detailEditRange.Columns[columnIndex++].Value = workData.IsOutOfExpiration ? DBConstant.ABSENT_EXCLUDED : work;
                detailEditRange.Columns[columnIndex++].Value = DBConstant.ABSENT_OFF.Equals(work) ? "" : workData.TimeTable;
                detailEditRange.Columns[columnIndex].Value = CommonUtil.ToDispHour(workData.Update_start_time);

                if (workData.IsLate)
                {
                    detailEditRange.Columns[columnIndex].Font.Color = CommonUtil.GetExcelColor(DBConstant.COLOR_RED);
                }
                columnIndex++;
                detailEditRange.Columns[columnIndex].Value = CommonUtil.ToDispHour(workData.Update_end_time);
                if (workData.IsLeaveEarly)
                {
                    detailEditRange.Columns[columnIndex].Font.Color = CommonUtil.GetExcelColor(DBConstant.COLOR_RED);
                }
                if (workData.IsOutOfExpiration)
                {
                    CommonUtil.SetBackColor(detailEditRange, 1, columnIndex, CommonUtil.GetExcelColor(DBConstant.COLOR_GRAY));
                }
                else if (workData.IsHoliday)
                {
                    CommonUtil.SetBackColor(detailEditRange, 1, columnIndex, CommonUtil.GetExcelColor(DBConstant.COLOR_HOLIDAY_XLS));
                }

                rowIndex++;

            }


            int totalColumn = 9;
            string presenceCol = "F";
            if (!isAllDepartment)
            {
                totalColumn--;
                presenceCol = "E";
            }
            // set total
            dailySheet.Cells[rowIndex++, totalColumn].Value = workDataList.Count;
            // set total present
            dailySheet.Cells[rowIndex++, totalColumn].Value = string.Format(COUNT_PRESENT, detailStartRowIndex, endRow, presenceCol, DBConstant.ABSENT_PERMISSION);
            // set total absent
            dailySheet.Cells[rowIndex++, totalColumn].Value = string.Format(COUNT_ABSENT, detailStartRowIndex, endRow, "{", "}",
                DBConstant.ABSENT_ANNUAL_LEAVE,
                DBConstant.ABSENT_SPECIAL_LEAVE,
                DBConstant.ABSENT_MATERNITY_LEAVE,
                DBConstant.ABSENT_PERMISSION,
                DBConstant.ABSENT_WITHOUT_PERMISSION,
                DBConstant.ABSENT_PERMISSION,
                DBConstant.ABSENT_OFF,
                presenceCol);
            // including
            rowIndex++;
            // OFF
            dailySheet.Cells[rowIndex++, totalColumn].Value = string.Format(COUNT_ABSENT_DETAIL, detailStartRowIndex, endRow, DBConstant.ABSENT_OFF, presenceCol);
            // set total absent al
            dailySheet.Cells[rowIndex++, totalColumn].Value = string.Format(COUNT_ABSENT_DETAIL, detailStartRowIndex, endRow, DBConstant.ABSENT_ANNUAL_LEAVE, presenceCol);
            // set total absent SL
            dailySheet.Cells[rowIndex++, totalColumn].Value = string.Format(COUNT_ABSENT_DETAIL, detailStartRowIndex, endRow, DBConstant.ABSENT_SPECIAL_LEAVE, presenceCol);
            // set total absent ML
            dailySheet.Cells[rowIndex++, totalColumn].Value = string.Format(COUNT_ABSENT_DETAIL, detailStartRowIndex, endRow, DBConstant.ABSENT_MATERNITY_LEAVE, presenceCol);
            // set total absent P
            dailySheet.Cells[rowIndex++, totalColumn].Value = string.Format(COUNT_ABSENT_DETAIL_P, detailStartRowIndex, endRow, DBConstant.ABSENT_PERMISSION, presenceCol);
            // set total absent WP
            dailySheet.Cells[rowIndex++, totalColumn].Value = string.Format(COUNT_ABSENT_DETAIL, detailStartRowIndex, endRow, DBConstant.ABSENT_WITHOUT_PERMISSION, presenceCol);
            // set total absent Excluded
            dailySheet.Cells[rowIndex++, totalColumn].Value = string.Format(COUNT_ABSENT_DETAIL, detailStartRowIndex, endRow, DBConstant.CHAR_UNKNOWN, presenceCol);
        }

        private void SetDutySheet(Worksheet template, Worksheet dutySheet, List<WorkDataModel> workDataList)
        {
            Range headerCompanyRange = dutySheet.get_Range("A2", Type.Missing);
            Range headerReportNameRange = dutySheet.get_Range("A3", Type.Missing);
            headerReportNameRange.Columns[1].Value = FormatHeaderByDay(headerReportNameRange.Columns[1].Value);
            headerCompanyRange.Columns[1].Value = selectedCompany.Company_name;

            Range detailRange = dutySheet.get_Range("A7", Type.Missing);
            int startRow = 7;
            string detailRowStart = "A{0}";
            int count = 0;
            int totalRow = workDataList.Count + startRow;
            if (totalRow == startRow)
            {
                totalRow++;
            }
            CommonUtil.CreateEmptyRow(dutySheet, startRow, workDataList.Count);

            foreach (var workData in workDataList)
            {
                int columnIndex = 1;

                Range detailEditRange = dutySheet.get_Range(string.Format(detailRowStart, startRow), Type.Missing);

                detailEditRange.Columns[columnIndex++].Value = ++count;
                detailEditRange.Columns[columnIndex++].Value = CommonUtil.GetXlsValue(workData.Post_name);
                detailEditRange.Columns[columnIndex++].Value = workData.EmployeeName;
                detailEditRange.Columns[columnIndex++].Value = workData.Employee_no;
                detailEditRange.Columns[columnIndex++].Value = CommonUtil.GetPosition(workData.Employee_remarks);
                detailEditRange.Columns[columnIndex++].Value = CommonUtil.ToDispHour(workData.Start_time, false);
                detailEditRange.Columns[columnIndex++].Value = CommonUtil.ToDispHour(workData.End_time, false);
                startRow++;

            }
            Range totalRange = dutySheet.Range["G" + totalRow, Type.Missing];
            totalRange.Columns[1].Value = workDataList.Count;
        }
        private void SetLateEarlySheet(Worksheet dutySheet, List<WorkDataModel> workDataList)
        {
            Range headerCompanyRange = dutySheet.get_Range("A2", Type.Missing);
            Range headerReportNameRange = dutySheet.get_Range("A3", Type.Missing);
            headerReportNameRange.Columns[1].Value = FormatHeaderByDay(headerReportNameRange.Columns[1].Value);
            headerCompanyRange.Columns[1].Value = selectedCompany.Company_name;
            Range detailRange = dutySheet.get_Range("A7", Type.Missing);
            int startRow = 7;
            string detailRowStart = "A{0}";
            int count = 0;
            int totalRow = workDataList.Count + startRow;
            if (totalRow == startRow)
            {
                totalRow++;
            }
            CommonUtil.CreateEmptyRow(dutySheet, startRow, workDataList.Count);
            foreach (var workData in workDataList)
            {
                int columnIndex = 1;

                Range detailEditRange = dutySheet.get_Range(string.Format(detailRowStart, startRow), Type.Missing);

                detailEditRange.Columns[columnIndex++].Value = ++count;
                detailEditRange.Columns[columnIndex++].Value = CommonUtil.GetXlsValue(workData.Post_name);
                detailEditRange.Columns[columnIndex++].Value = workData.EmployeeName;
                detailEditRange.Columns[columnIndex++].Value = workData.Employee_no;
                detailEditRange.Columns[columnIndex++].Value = CommonUtil.GetPosition(workData.Employee_remarks);
                detailEditRange.Columns[columnIndex++].Value = workData.TimeTable;
                detailEditRange.Columns[columnIndex++].Value = CommonUtil.ToDispHour(workData.Update_start_time, false);
                detailEditRange.Columns[columnIndex++].Value = CommonUtil.ToDispHour(workData.Update_end_time, false);
                startRow++;

            }
            Range totalRange = dutySheet.Range["H" + totalRow, Type.Missing];
            totalRange.Columns[1].Value = workDataList.Count;
        }
        private void SetReCAP1Detail(Worksheet recapSheet, decimal postNo, int rowIndex, List<WorkDataModel> workDataList)
        {
            string detailRowStart = "A{0}";
            int columnIndex = 1;
            decimal absent = 0;
            decimal present = 0;
            int off = 0;
            int al = 0;
            int sl = 0;
            int ml = 0;
            decimal p = 0;
            int wp = 0;
            int noOnduty = 0;
            int cameLate = 0;
            int noOffDuty = 0;
            int leftEarly = 0;
            decimal halfP = 0;
            List<WorkDataModel> workDataByPost = workDataList.FindAll(x => x.Post_no == postNo);
            Range detailEditRange = recapSheet.get_Range(string.Format(detailRowStart, rowIndex), Type.Missing);

            present = workDataByPost.Count(x => x.WorkingType == WorkingType.Present);
            absent = workDataByPost.Count(x => x.WorkingType == WorkingType.Absent && x.Work_type_no != DBConstant.WORK_TYPE_HALF_PERMISSION);
            al = workDataByPost.Count(x => !x.IsHoliday && (x.Work_type_no == DBConstant.WORK_TYPE_ANNUAL_LEAVE));
            sl = workDataByPost.Count(x => !x.IsHoliday && (x.Work_type_no == DBConstant.WORK_TYPE_SPECIAL_LEAVE));
            ml = workDataByPost.Count(x => !x.IsHoliday && (x.Work_type_no == DBConstant.WORK_TYPE_MATERNITY_LEAVE));
            p = workDataByPost.Count(x => !x.IsHoliday && (x.Work_type_no == DBConstant.WORK_TYPE_PERMISSION));
            halfP = workDataByPost.Count(x => x.Work_type_no == DBConstant.WORK_TYPE_HALF_PERMISSION);
            wp = workDataByPost.Count(x => !x.IsHoliday && (x.Work_type_no == DBConstant.WORK_TYPE_AW_PERMISSION));
            noOnduty = workDataByPost.Count(x => x.IsNoOnDuty);
            noOffDuty = workDataByPost.Count(x => x.IsNoOffDuty);
            cameLate = workDataByPost.Count(x => x.IsLate);
            leftEarly = workDataByPost.Count(x => x.IsLeaveEarly);
            off = workDataByPost.Count(x => x.WorkingType == WorkingType.Absent && x.IsHoliday);
            halfP = halfP / 2;
            absent += halfP;
            present -= halfP;
            p += halfP;

            detailEditRange.Columns[columnIndex++].Value = CommonUtil.GetXlsValue(workDataByPost.FirstOrDefault().Post_name, false);
            detailEditRange.Columns[columnIndex++].Value = CommonUtil.GetXlsValue(workDataByPost.Count);
            detailEditRange.Columns[columnIndex++].Value = CommonUtil.GetXlsValue(present);
            detailEditRange.Columns[columnIndex++].Value = CommonUtil.GetXlsValue(absent);
            detailEditRange.Columns[columnIndex++].Value = CommonUtil.GetXlsValue(off);
            detailEditRange.Columns[columnIndex++].Value = CommonUtil.GetXlsValue(al);
            detailEditRange.Columns[columnIndex++].Value = CommonUtil.GetXlsValue(sl);
            detailEditRange.Columns[columnIndex++].Value = CommonUtil.GetXlsValue(ml);
            detailEditRange.Columns[columnIndex++].Value = CommonUtil.GetXlsValue(p);
            detailEditRange.Columns[columnIndex++].Value = CommonUtil.GetXlsValue(wp);
            detailEditRange.Columns[columnIndex++].Value = CommonUtil.GetXlsValue(noOnduty);
            detailEditRange.Columns[columnIndex++].Value = CommonUtil.GetXlsValue(cameLate);
            detailEditRange.Columns[columnIndex++].Value = CommonUtil.GetXlsValue(noOffDuty);
            detailEditRange.Columns[columnIndex++].Value = CommonUtil.GetXlsValue(leftEarly);

        }

        private void SetRECAP1(Worksheet template, Worksheet recapSheet, List<WorkDataModel> workDataList, bool isKnitting)
        {
            Range headerCompanyRange = recapSheet.get_Range("A2", Type.Missing);
            Range headerReportNameRange = recapSheet.get_Range("A3", Type.Missing);
            headerReportNameRange.Columns[1].Value = FormatHeaderByDay(headerReportNameRange.Columns[1].Value);
            headerCompanyRange.Columns[1].Value = selectedCompany.Company_name;
            Range detailRange = recapSheet.get_Range("A7", Type.Missing);
            int startNormalRow = 7;
            int rowIndex = startNormalRow;

            if (isKnitting)
            {
                SortDataRecap1Knitting(workDataList);
            }
            else
            {
                SortDataByShift(workDataList);
            }
            var groupbyDepartsByTeam = workDataList.GroupBy(x => x.Post_no);

            Dictionary<decimal, List<WorkDataModel>> normalWorkingMap = new Dictionary<decimal, List<WorkDataModel>>();
            Dictionary<decimal, List<WorkDataModel>> offMap = new Dictionary<decimal, List<WorkDataModel>>();

            foreach (var post in groupbyDepartsByTeam)
            {
                List<WorkDataModel> workDataByPost = workDataList.FindAll(x => x.Post_no == post.Key);
                if (workDataByPost.Count(x => x.IsHoliday) > workDataByPost.Count / 2)
                {
                    offMap.Add(post.Key, workDataByPost);
                }
                else
                {
                    normalWorkingMap.Add(post.Key, workDataByPost);
                }
            }

            int totalNormalRow = normalWorkingMap.Keys.Count + rowIndex;
            int startOffRow = totalNormalRow + 1;
            int totalColumn = 2;
            int endRow = rowIndex;
            int startOffRowIndex;

            string sum = "=SUM({0}{1}:{0}{2})";
            Range totalRange;

            CommonUtil.CreateEmptyRow(recapSheet, startNormalRow, normalWorkingMap.Keys.Count);
            CommonUtil.CreateEmptyRow(recapSheet, totalNormalRow + 1, offMap.Keys.Count);
            if (normalWorkingMap.Keys.Count > 0)
            {
                foreach (var key in normalWorkingMap.Keys)
                {
                    SetReCAP1Detail(recapSheet, key, rowIndex++, workDataList);
                }

            }
            else
            {
                rowIndex++;
            }
            endRow = rowIndex - 1;
            startOffRowIndex = rowIndex;
            totalRange = recapSheet.get_Range("A" + rowIndex, "N" + rowIndex);


            //Total
            totalRange.Columns[totalColumn].Value = string.Format(sum, CommonUtil.GetExcelColumnName(totalColumn++), startNormalRow, endRow);
            totalRange.Columns[totalColumn].Value = string.Format(sum, CommonUtil.GetExcelColumnName(totalColumn++), startNormalRow, endRow);
            totalRange.Columns[totalColumn].Value = string.Format(sum, CommonUtil.GetExcelColumnName(totalColumn++), startNormalRow, endRow);
            totalRange.Columns[totalColumn].Value = string.Format(sum, CommonUtil.GetExcelColumnName(totalColumn++), startNormalRow, endRow);
            totalRange.Columns[totalColumn].Value = string.Format(sum, CommonUtil.GetExcelColumnName(totalColumn++), startNormalRow, endRow);
            totalRange.Columns[totalColumn].Value = string.Format(sum, CommonUtil.GetExcelColumnName(totalColumn++), startNormalRow, endRow);
            totalRange.Columns[totalColumn].Value = string.Format(sum, CommonUtil.GetExcelColumnName(totalColumn++), startNormalRow, endRow);
            totalRange.Columns[totalColumn].Value = string.Format(sum, CommonUtil.GetExcelColumnName(totalColumn++), startNormalRow, endRow);
            totalRange.Columns[totalColumn].Value = string.Format(sum, CommonUtil.GetExcelColumnName(totalColumn++), startNormalRow, endRow);
            totalRange.Columns[totalColumn].Value = string.Format(sum, CommonUtil.GetExcelColumnName(totalColumn++), startNormalRow, endRow);
            totalRange.Columns[totalColumn].Value = string.Format(sum, CommonUtil.GetExcelColumnName(totalColumn++), startNormalRow, endRow);
            totalRange.Columns[totalColumn].Value = string.Format(sum, CommonUtil.GetExcelColumnName(totalColumn++), startNormalRow, endRow);
            totalRange.Columns[totalColumn].Value = string.Format(sum, CommonUtil.GetExcelColumnName(totalColumn++), startNormalRow, endRow);
            //off
            rowIndex++;
            if (offMap.Keys.Count > 0)
            {
                foreach (var key in offMap.Keys)
                {
                    SetReCAP1Detail(recapSheet, key, rowIndex++, workDataList);
                }

            }
            else
            {
                rowIndex++;

            }
            endRow = rowIndex - 1;
            // grand total
            totalRange = recapSheet.get_Range("A" + rowIndex, "N" + rowIndex);

            totalColumn = 2;
            totalRange.Columns[totalColumn].Value = string.Format(sum, CommonUtil.GetExcelColumnName(totalColumn++), startOffRowIndex, endRow);
            totalRange.Columns[totalColumn].Value = string.Format(sum, CommonUtil.GetExcelColumnName(totalColumn++), startOffRowIndex, endRow);
            totalRange.Columns[totalColumn].Value = string.Format(sum, CommonUtil.GetExcelColumnName(totalColumn++), startOffRowIndex, endRow);
            totalRange.Columns[totalColumn].Value = string.Format(sum, CommonUtil.GetExcelColumnName(totalColumn++), startOffRowIndex, endRow);
            totalRange.Columns[totalColumn].Value = string.Format(sum, CommonUtil.GetExcelColumnName(totalColumn++), startOffRowIndex, endRow);
            totalRange.Columns[totalColumn].Value = string.Format(sum, CommonUtil.GetExcelColumnName(totalColumn++), startOffRowIndex, endRow);
            totalRange.Columns[totalColumn].Value = string.Format(sum, CommonUtil.GetExcelColumnName(totalColumn++), startOffRowIndex, endRow);
            totalRange.Columns[totalColumn].Value = string.Format(sum, CommonUtil.GetExcelColumnName(totalColumn++), startOffRowIndex, endRow);
            totalRange.Columns[totalColumn].Value = string.Format(sum, CommonUtil.GetExcelColumnName(totalColumn++), startOffRowIndex, endRow);
            totalRange.Columns[totalColumn].Value = string.Format(sum, CommonUtil.GetExcelColumnName(totalColumn++), startOffRowIndex, endRow);
            totalRange.Columns[totalColumn].Value = string.Format(sum, CommonUtil.GetExcelColumnName(totalColumn++), startOffRowIndex, endRow);
            totalRange.Columns[totalColumn].Value = string.Format(sum, CommonUtil.GetExcelColumnName(totalColumn++), startOffRowIndex, endRow);
            totalRange.Columns[totalColumn].Value = string.Format(sum, CommonUtil.GetExcelColumnName(totalColumn++), startOffRowIndex, endRow);

        }
        private void SetRECAPConfectionOrExpedition(Worksheet template, Worksheet recapSheet, List<WorkDataModel> workDataList)
        {
            Range headerCompanyRange = recapSheet.get_Range("A2", Type.Missing);
            Range headerReportNameRange = recapSheet.get_Range("A3", Type.Missing);
            headerReportNameRange.Columns[1].Value = FormatHeaderByDay(headerReportNameRange.Columns[1].Value);
            headerCompanyRange.Columns[1].Value = selectedCompany.Company_name;
            Range detailRange = recapSheet.get_Range("A7", Type.Missing);
            int startRow = 7;
            int rowIndex = startRow;
            string detailRowStart = "A{0}";

            var groupbyDeparts = workDataList.GroupBy(x => x.Position).OrderBy(y => y.Key);
            int endRow = startRow + groupbyDeparts.Count() - 1;
            CommonUtil.CreateEmptyRow(recapSheet, startRow, groupbyDeparts.Count());
            if (endRow < startRow)
            {
                endRow++;
            }
            foreach (var post in groupbyDeparts)
            {
                int columnIndex = 1;
                decimal present = 0;
                decimal absent = 0;
                decimal halfP = 0;
                decimal off = 0;
                int al = 0;
                int sl = 0;
                int ml = 0;
                decimal p = 0;
                int wp = 0;
                int noOnduty = 0;
                int cameLate = 0;
                int noOffDuty = 0;
                int leftEarly = 0;

                List<WorkDataModel> workDataByPosition = workDataList.FindAll(x => x.Position == post.Key);
                Range detailEditRange = recapSheet.get_Range(string.Format(detailRowStart, rowIndex), Type.Missing);

                present = workDataByPosition.Count(x => x.WorkingType == WorkingType.Present);
                absent = workDataByPosition.Count(x => x.WorkingType == WorkingType.Absent && x.Work_type_no != DBConstant.WORK_TYPE_HALF_PERMISSION);
                halfP = workDataByPosition.Count(x => x.Work_type_no == DBConstant.WORK_TYPE_HALF_PERMISSION);
                al = workDataByPosition.Count(x => (x.Work_type_no == DBConstant.WORK_TYPE_ANNUAL_LEAVE) && !x.IsHoliday);
                sl = workDataByPosition.Count(x => (x.Work_type_no == DBConstant.WORK_TYPE_SPECIAL_LEAVE) && !x.IsHoliday);
                ml = workDataByPosition.Count(x => (x.Work_type_no == DBConstant.WORK_TYPE_MATERNITY_LEAVE) && !x.IsHoliday);
                p = workDataByPosition.Count(x => (x.Work_type_no == DBConstant.WORK_TYPE_PERMISSION) && !x.IsHoliday);
                wp = workDataByPosition.Count(x => (x.Work_type_no == DBConstant.WORK_TYPE_AW_PERMISSION) && !x.IsHoliday);
                noOnduty = workDataByPosition.Count(x => x.IsNoOnDuty);
                noOffDuty = workDataByPosition.Count(x => x.IsNoOffDuty);
                cameLate = workDataByPosition.Count(x => x.IsLate);
                leftEarly = workDataByPosition.Count(x => x.IsLeaveEarly);
                off = workDataByPosition.Count(x => x.WorkingType == WorkingType.Absent && x.IsHoliday);
                halfP = halfP / 2;
                absent += halfP;
                p += halfP;
                present -= halfP;
                detailEditRange.Columns[columnIndex++].Value = post.Key;
                detailEditRange.Columns[columnIndex++].Value = CommonUtil.GetXlsValue(workDataByPosition.Count);
                detailEditRange.Columns[columnIndex++].Value = CommonUtil.GetXlsValue(present);
                detailEditRange.Columns[columnIndex++].Value = CommonUtil.GetXlsValue(absent);
                detailEditRange.Columns[columnIndex++].Value = CommonUtil.GetXlsValue(off);
                detailEditRange.Columns[columnIndex++].Value = CommonUtil.GetXlsValue(al);
                detailEditRange.Columns[columnIndex++].Value = CommonUtil.GetXlsValue(sl);
                detailEditRange.Columns[columnIndex++].Value = CommonUtil.GetXlsValue(ml);
                detailEditRange.Columns[columnIndex++].Value = CommonUtil.GetXlsValue(p);
                detailEditRange.Columns[columnIndex++].Value = CommonUtil.GetXlsValue(wp);
                rowIndex++;

            }

            int totalColumn = 2;

            string sum = "=SUM({0}{1}:{0}{2})";

            Range totalRange = recapSheet.get_Range("A" + (endRow + 1), Type.Missing);
            //Total
            totalRange.Columns[totalColumn].Value = string.Format(sum, CommonUtil.GetExcelColumnName(totalColumn++), startRow, endRow);
            totalRange.Columns[totalColumn].Value = string.Format(sum, CommonUtil.GetExcelColumnName(totalColumn++), startRow, endRow);
            totalRange.Columns[totalColumn].Value = string.Format(sum, CommonUtil.GetExcelColumnName(totalColumn++), startRow, endRow);
            totalRange.Columns[totalColumn].Value = string.Format(sum, CommonUtil.GetExcelColumnName(totalColumn++), startRow, endRow);
            totalRange.Columns[totalColumn].Value = string.Format(sum, CommonUtil.GetExcelColumnName(totalColumn++), startRow, endRow);
            totalRange.Columns[totalColumn].Value = string.Format(sum, CommonUtil.GetExcelColumnName(totalColumn++), startRow, endRow);
            totalRange.Columns[totalColumn].Value = string.Format(sum, CommonUtil.GetExcelColumnName(totalColumn++), startRow, endRow);
            totalRange.Columns[totalColumn].Value = string.Format(sum, CommonUtil.GetExcelColumnName(totalColumn++), startRow, endRow);
            totalRange.Columns[totalColumn].Value = string.Format(sum, CommonUtil.GetExcelColumnName(totalColumn++), startRow, endRow);

        }

        private void SortDataByShiftForDaily(List<WorkDataModel> workDataList)
        {

            workDataList.Sort(delegate(WorkDataModel o1, WorkDataModel o2)
            {

                string start_time_o1 = string.IsNullOrEmpty(o1.Update_start_time) ? string.Empty : o1.Update_start_time.PadLeft(4);
                string start_time_o2 = string.IsNullOrEmpty(o2.Update_start_time) ? string.Empty : o2.Update_start_time.PadLeft(4);
                int timeCompare = timeCompare = o1.Post_name.CompareTo(o2.Post_name);
                if (timeCompare == 0)
                {

                    timeCompare = start_time_o1.CompareTo(start_time_o2);

                }
                return timeCompare;
            });


        }

        private void SortDataRecap1Knitting(List<WorkDataModel> workDataList)
        {

            SortDataByShiftForDaily(workDataList);

            List<WorkDataModel> knittingTeams = workDataList.FindAll(x => x.Post_name.ToUpper().StartsWith(DBConstant.RECAP_2_START_WITH_KNITTING_DEPRT));
            int kinttingIndex = workDataList.FindIndex(x => x.Post_name.ToUpper().StartsWith(DBConstant.RECAP_2_START_WITH_KNITTING_DEPRT));
            if (kinttingIndex >= 0)
            {
                var regular = knittingTeams.FindAll(x => !x.Post_name.ToUpper().Contains("TEAM"));
                var teams = knittingTeams.FindAll(x => x.Post_name.ToUpper().Contains("TEAM"));
                List<WorkDataModel> workByTeams = new List<WorkDataModel>();

                List<WorkDataModel> workTempList = new List<WorkDataModel>();

                var groupByPostNo = teams.GroupBy(x => x.Post_no).Select(x => x.Key).ToList();

                foreach (var postNo in groupByPostNo)
                {
                    var team = teams.FindAll(x => x.Post_no == postNo);
                    var groupByTime = team.GroupBy(x => x.Work_from).Select(x => x.Key).ToList();
                    var maxrecord = 0;
                    decimal workFrom = 0;
                    foreach (var workTime in groupByTime)
                    {
                        var countByTime = team.Count(x => x.Work_from == workTime);
                        if (countByTime > maxrecord)
                        {
                            maxrecord = countByTime;
                            workFrom = workTime;
                        }
                    }
                    workTempList.Add(team.Find(x => x.Work_from == workFrom));
                }

                workTempList.Sort(delegate(WorkDataModel o1, WorkDataModel o2)
                {
                    return o1.Work_from.CompareTo(o2.Work_from);
                });

                foreach (var workData in workTempList)
                {
                    workByTeams.AddRange(teams.FindAll(x => x.Post_no == workData.Post_no));
                }
                knittingTeams.Clear();
                knittingTeams.AddRange(regular);
                knittingTeams.AddRange(workByTeams);

                workDataList.RemoveRange(kinttingIndex, knittingTeams.Count);
                workDataList.InsertRange(kinttingIndex, knittingTeams);
            }

        }
        private void SortDataByShift(List<WorkDataModel> workDataList)
        {
            List<WorkDataModel> regularDeparts = workDataList.FindAll(x => !x.Post_name.ToLower().Contains("team"));
            List<WorkDataModel> departByTeams = workDataList.FindAll(x => x.Post_name.ToLower().Contains("team"));
            List<WorkDataModel> workByTeams = new List<WorkDataModel>();
            List<WorkDataModel> workTempList = new List<WorkDataModel>();

            var groupByPostNo = departByTeams.GroupBy(x => x.Post_no).Select(x => x.Key).ToList();

            foreach (var postNo in groupByPostNo)
            {
                var team = departByTeams.FindAll(x => x.Post_no == postNo);
                var groupByTime = team.GroupBy(x => x.Work_from).Select(x => x.Key).ToList();
                var maxrecord = 0;
                decimal workFrom = 0;
                foreach (var workTime in groupByTime)
                {
                    var countByTime = team.Count(x => x.Work_from == workTime);
                    if (countByTime > maxrecord)
                    {
                        maxrecord = countByTime;
                        workFrom = workTime;
                    }
                }
                workTempList.Add(team.Find(x => x.Work_from == workFrom));
            }

            workTempList.Sort(delegate(WorkDataModel o1, WorkDataModel o2)
            {
                return o1.Work_from.CompareTo(o2.Work_from);
            });

            foreach (var workData in workTempList)
            {
                workByTeams.AddRange(departByTeams.FindAll(x => x.Post_no == workData.Post_no));
            }
            workDataList.Clear();
            workDataList.AddRange(regularDeparts);
            workDataList.AddRange(workByTeams);
        }
        private void SetRECAP2(Worksheet template, Worksheet recapSheet, List<WorkDataModel> workDataList, bool isSpinning)
        {
            Range headerCompanyRange = recapSheet.get_Range("A2", Type.Missing);
            Range headerReportNameRange = recapSheet.get_Range("A3", Type.Missing);
            Range headerReportTypeNameRange = recapSheet.get_Range("A4", Type.Missing);
            headerReportNameRange.Columns[1].Value = FormatHeaderByDay(headerReportNameRange.Columns[1].Value);
            headerCompanyRange.Columns[1].Value = selectedCompany.Company_name;
            if (isSpinning)
            {
                headerReportTypeNameRange.Columns[1].Value = DBConstant.SHEET_RECAP_SPINNING_2_HEADER;
            }
            else
            {
                headerReportTypeNameRange.Columns[1].Value = DBConstant.SHEET_RECAP_KNITTING_2_HEADER;
            }
            SortDataByShift(workDataList);
            List<WorkDataModel> regularDeparts = workDataList.FindAll(x => !x.Post_name.ToLower().Contains("team"));
            List<WorkDataModel> departByTeams = workDataList.FindAll(x => x.Post_name.ToLower().Contains("team"));
            Dictionary<decimal, List<WorkDataModel>> regularMap = new Dictionary<decimal, List<WorkDataModel>>();
            Dictionary<decimal, List<WorkDataModel>> teamsMap = new Dictionary<decimal, List<WorkDataModel>>();
            Dictionary<decimal, List<WorkDataModel>> offMap = new Dictionary<decimal, List<WorkDataModel>>();
            int totalDepartment = 0;
            int holidayColor = CommonUtil.GetExcelColor(DBConstant.COLOR_HOLIDAY_XLS);

            var groupByDepartments = departByTeams.GroupBy(x => x.Post_no);
            // ordered by time_no
            foreach (var group in groupByDepartments)
            {
                var workListByDepartments = departByTeams.FindAll(x => x.Post_no == group.Key);
                // if most are holiday then set off for deparment
                if (workListByDepartments.Count(x => x.IsHoliday) > workListByDepartments.Count / 2)
                {
                    offMap.Add(group.Key, workListByDepartments);
                }
                else
                {
                    teamsMap.Add(group.Key, workListByDepartments);
                }
                totalDepartment++;
            }

            foreach (var work in regularDeparts)
            {
                List<WorkDataModel> workList;
                if (!regularMap.ContainsKey(work.Post_no))
                {
                    workList = new List<WorkDataModel>();
                    regularMap.Add(work.Post_no, workList);
                    totalDepartment++;
                }
                else
                {
                    workList = regularMap[work.Post_no];
                }
                workList.Add(work);
            }

            #region create header

            for (int i = 0; i < totalDepartment - 1; i++)
            {
                Range range = recapSheet.Range["B1", "D1"].EntireColumn;
                range.Copy(Type.Missing);
                range.Insert(XlInsertShiftDirection.xlShiftToRight, Type.Missing);
            }
            int columnIndex = 2;
            int rowPostName = 6;
            int rowTimeTable = 7;
            foreach (var key in regularMap.Keys)
            {
                WorkDataModel workByPost = regularMap[key].FindAll(x => x.WorkingType == WorkingType.Present).FirstOrDefault();

                if (workByPost != null)
                {
                    recapSheet.Cells[rowPostName, columnIndex].Value = workByPost.Post_name;
                    recapSheet.Cells[rowTimeTable, columnIndex].Value = workByPost.TimeTable;
                }
                else
                {
                    workByPost = regularMap[key].FirstOrDefault();
                    recapSheet.Cells[rowPostName, columnIndex].Value = workByPost.Post_name;
                    recapSheet.Cells[rowTimeTable, columnIndex].Value = DBConstant.RECAP_2_NOT_SAME_SHIFT_DISP;
                }
                columnIndex += 3;

            }


            foreach (var key in teamsMap.Keys)
            {
                PostModel post = departmentList.Find(x => x.Post_no == key.ToString());
                List<WorkDataModel> workDataByTeam = teamsMap[key];
                List<WorkDataModel> workDataByTeamPresent = workDataByTeam.FindAll(x => x.WorkingType != WorkingType.Absent);

                recapSheet.Cells[rowPostName, columnIndex].Value = post.Post_name;
                int shiftCount = workDataByTeamPresent.GroupBy(x => x.Time_table_no).Count();
                if (shiftCount > 0)
                {
                    if (shiftCount != 1)
                    {
                        recapSheet.Cells[rowTimeTable, columnIndex].Value = DBConstant.RECAP_2_NOT_SAME_SHIFT_DISP;
                    }
                    else
                    {
                        WorkDataModel workByPost = workDataByTeamPresent.FirstOrDefault();

                        recapSheet.Cells[rowTimeTable, columnIndex].Value = workByPost.TimeTable;

                    }
                }
                else
                {
                    recapSheet.Cells[rowTimeTable, columnIndex].Value = workDataByTeam.FirstOrDefault().TimeTable;
                }
                columnIndex += 3;

            }
            foreach (var key in offMap.Keys)
            {
                PostModel post = departmentList.Find(x => x.Post_no == key.ToString());
                List<WorkDataModel> workDataByTeamOff = offMap[key];

                WorkDataModel workByPost = workDataByTeamOff.FirstOrDefault();


                Range postHeader = recapSheet.Cells[rowPostName, columnIndex];
                Range postTimeTable = recapSheet.Cells[rowTimeTable, columnIndex];


                Range topTotalRange = recapSheet.Cells[rowTimeTable + 1, columnIndex];
                Range topOnRange = recapSheet.Cells[rowTimeTable + 1, columnIndex + 1];
                Range topOffRange = recapSheet.Cells[rowTimeTable + 1, columnIndex + 2];

                Range totalTotalRange = recapSheet.Cells[rowTimeTable + 2, columnIndex];
                Range totalOnRange = recapSheet.Cells[rowTimeTable + 2, columnIndex + 1];
                Range totalOffRange = recapSheet.Cells[rowTimeTable + 2, columnIndex + 2];

                Range bottomTotalRange = recapSheet.Cells[rowTimeTable + 3, columnIndex];
                Range bottomOnRange = recapSheet.Cells[rowTimeTable + 3, columnIndex + 1];
                Range bottomOffRange = recapSheet.Cells[rowTimeTable + 3, columnIndex + 2];

                postHeader.Value = post.Post_name;
                postTimeTable.Value = DBConstant.ABSENT_OFF;
                postHeader.Interior.Color = holidayColor;
                postTimeTable.Interior.Color = holidayColor;

                topTotalRange.Interior.Color = holidayColor;
                topOnRange.Interior.Color = holidayColor;
                topOffRange.Interior.Color = holidayColor;

                totalTotalRange.Interior.Color = holidayColor;
                totalOnRange.Interior.Color = holidayColor;
                totalOffRange.Interior.Color = holidayColor;

                bottomTotalRange.Interior.Color = holidayColor;
                bottomOnRange.Interior.Color = holidayColor;
                bottomOffRange.Interior.Color = holidayColor;
                columnIndex += 3;

            }
            #endregion
            #region detail
            int startRow = 9;
            int rowIndex = startRow;
            string detailRowStart = "A{0}";
            var groupByPositions = workDataList.GroupBy(x => x.Position).OrderBy(y => y.Key);
            int totalRow = groupByPositions.Count() + rowIndex;
            CommonUtil.CreateEmptyRow(recapSheet, startRow, groupByPositions.Count());
            foreach (var post in groupByPositions)
            {
                Range detailEditRange = recapSheet.get_Range(string.Format(detailRowStart, rowIndex), Type.Missing);
                string totalEmployee = "=";
                string totalEmployeeOn = "=";
                string totalEmployeeOff = "=";
                string totalHolidayOff = "=";
                columnIndex = 1;
                int total = 0;
                decimal totalPresence = 0;
                decimal totalAbsent = 0;
                decimal present = 0;
                decimal absent = 0;
                decimal halfP = 0;
                int al = 0;
                int sl = 0;
                int ml = 0;
                decimal p = 0;
                int wp = 0;
                decimal off = 0;

                detailEditRange.Columns[columnIndex++].Value = post.Key;
                if (string.IsNullOrEmpty(post.Key) || string.IsNullOrWhiteSpace(post.Key))
                {
                    string test = post.Key;
                }
                foreach (var key in regularMap.Keys)
                {
                    List<WorkDataModel> workDataByRegularShift = regularMap[key].FindAll(x => x.Position == post.Key && x.Post_no == key);
                    if (workDataByRegularShift.Count > 0)
                    {
                        present = workDataByRegularShift.Count(x => x.WorkingType == WorkingType.Present);
                        absent = workDataByRegularShift.Count(x => x.WorkingType == WorkingType.Absent && x.Work_type_no != DBConstant.WORK_TYPE_HALF_PERMISSION);

                        halfP = workDataByRegularShift.Count(x => x.Work_type_no == DBConstant.WORK_TYPE_HALF_PERMISSION);
                        al += workDataByRegularShift.Count(x => x.Work_type_no == DBConstant.WORK_TYPE_ANNUAL_LEAVE);
                        sl += workDataByRegularShift.Count(x => x.Work_type_no == DBConstant.WORK_TYPE_SPECIAL_LEAVE);
                        ml += workDataByRegularShift.Count(x => x.Work_type_no == DBConstant.WORK_TYPE_MATERNITY_LEAVE);
                        p += workDataByRegularShift.Count(x => x.Work_type_no == DBConstant.WORK_TYPE_PERMISSION);
                        wp += workDataByRegularShift.Count(x => x.Work_type_no == DBConstant.WORK_TYPE_AW_PERMISSION);
                        off += workDataByRegularShift.Count(x => x.WorkingType == WorkingType.Absent && x.IsHoliday);
                        totalPresence += present;
                        totalAbsent += absent;
                        total += workDataByRegularShift.Count;
                        halfP = halfP / 2;
                        absent += halfP;
                        p += halfP;
                        present -= halfP;
                        totalEmployee += CommonUtil.GetExcelColumnName(columnIndex) + rowIndex + "+";
                        detailEditRange.Columns[columnIndex++].Value = CommonUtil.GetXlsValue(workDataByRegularShift.Count);
                        totalEmployeeOn += CommonUtil.GetExcelColumnName(columnIndex) + rowIndex + "+";
                        detailEditRange.Columns[columnIndex++].Value = CommonUtil.GetXlsValue(present);
                        totalEmployeeOff += CommonUtil.GetExcelColumnName(columnIndex) + rowIndex + "+";
                        detailEditRange.Columns[columnIndex++].Value = CommonUtil.GetXlsValue(absent);
                    }
                    else
                    {
                        totalEmployee += CommonUtil.GetExcelColumnName(columnIndex++) + rowIndex + "+";
                        totalEmployeeOn += CommonUtil.GetExcelColumnName(columnIndex++) + rowIndex + "+";
                        totalEmployeeOff += CommonUtil.GetExcelColumnName(columnIndex++) + rowIndex + "+";
                    }
                }

                foreach (var key in teamsMap.Keys)
                {
                    List<WorkDataModel> workDataByTeam = teamsMap[key].FindAll(x => x.Position == post.Key && x.Post_no == key);

                    if (workDataByTeam.Count > 0)
                    {
                        present = workDataByTeam.Count(x => x.WorkingType == WorkingType.Present);
                        absent = workDataByTeam.Count(x => x.WorkingType == WorkingType.Absent);
                        halfP = workDataByTeam.Count(x => x.Work_type_no == DBConstant.WORK_TYPE_HALF_PERMISSION);

                        al += workDataByTeam.Count(x => x.Work_type_no == DBConstant.WORK_TYPE_ANNUAL_LEAVE);
                        sl += workDataByTeam.Count(x => x.Work_type_no == DBConstant.WORK_TYPE_SPECIAL_LEAVE);
                        ml += workDataByTeam.Count(x => x.Work_type_no == DBConstant.WORK_TYPE_MATERNITY_LEAVE);
                        p += workDataByTeam.Count(x => x.Work_type_no == DBConstant.WORK_TYPE_PERMISSION);
                        wp += workDataByTeam.Count(x => x.Work_type_no == DBConstant.WORK_TYPE_AW_PERMISSION);
                        off += workDataByTeam.Count(x => x.WorkingType == WorkingType.Absent && x.IsHoliday);

                        halfP = halfP / 2;
                        absent += halfP;
                        p += halfP;
                        present -= halfP;
                        totalPresence += present;
                        totalAbsent += absent;
                        total += workDataByTeam.Count;

                        totalEmployee += CommonUtil.GetExcelColumnName(columnIndex) + rowIndex + "+";
                        detailEditRange.Columns[columnIndex++].Value = CommonUtil.GetXlsValue(workDataByTeam.Count);
                        totalEmployeeOn += CommonUtil.GetExcelColumnName(columnIndex) + rowIndex + "+";
                        detailEditRange.Columns[columnIndex++].Value = CommonUtil.GetXlsValue(present);
                        totalEmployeeOff += CommonUtil.GetExcelColumnName(columnIndex) + rowIndex + "+";
                        detailEditRange.Columns[columnIndex++].Value = CommonUtil.GetXlsValue(absent);
                        //columnIndex++;
                    }
                    else
                    {
                        totalEmployee += CommonUtil.GetExcelColumnName(columnIndex++) + rowIndex + "+";
                        totalEmployeeOn += CommonUtil.GetExcelColumnName(columnIndex++) + rowIndex + "+";
                        totalEmployeeOff += CommonUtil.GetExcelColumnName(columnIndex++) + rowIndex + "+";
                    }

                }


                foreach (var key in offMap.Keys)
                {
                    List<WorkDataModel> workDataByOffTeam = offMap[key].FindAll(x => x.Position == post.Key && x.Post_no == key);
                    totalEmployee += CommonUtil.GetExcelColumnName(columnIndex) + rowIndex + "+";
                    Range totalRange = detailEditRange.Columns[columnIndex++];
                    totalEmployeeOn += CommonUtil.GetExcelColumnName(columnIndex) + rowIndex + "+";
                    Range presentRange = detailEditRange.Columns[columnIndex++];
                    totalEmployeeOff += CommonUtil.GetExcelColumnName(columnIndex) + rowIndex + "+";
                    totalHolidayOff += CommonUtil.GetExcelColumnName(columnIndex) + rowIndex + "+";
                    Range absentRange = detailEditRange.Columns[columnIndex++];

                    if (workDataByOffTeam.Count > 0)
                    {
                        present = workDataByOffTeam.Count(x => x.WorkingType == WorkingType.Present);
                        absent = workDataByOffTeam.Count(x => x.WorkingType == WorkingType.Absent);
                        totalRange.Value = CommonUtil.GetXlsValue(workDataByOffTeam.Count);
                        presentRange.Value = CommonUtil.GetXlsValue(present);
                        absentRange.Value = CommonUtil.GetXlsValue(absent);
                        off += absent;
                    }
                }

                //TOTAL

                detailEditRange.Columns[columnIndex++].Value = totalEmployee == "=" ? "" : totalEmployee.TrimEnd('+');
                detailEditRange.Columns[columnIndex++].Value = totalEmployeeOn == "=" ? "" : totalEmployeeOn.TrimEnd('+');
                detailEditRange.Columns[columnIndex++].Value = totalEmployeeOff == "=" ? "" : totalEmployeeOff.TrimEnd('+');
                //OFF
                detailEditRange.Columns[columnIndex++].Value = CommonUtil.GetXlsValue(off);
                detailEditRange.Columns[columnIndex++].Value = CommonUtil.GetXlsValue(al);
                detailEditRange.Columns[columnIndex++].Value = CommonUtil.GetXlsValue(sl);
                detailEditRange.Columns[columnIndex++].Value = CommonUtil.GetXlsValue(ml);
                detailEditRange.Columns[columnIndex++].Value = CommonUtil.GetXlsValue(p);
                detailEditRange.Columns[columnIndex++].Value = CommonUtil.GetXlsValue(wp);
                rowIndex++;
            }

            int startColumnDetail = 2;
            int endRow = rowIndex - 1;
            if (totalRow == startRow)
            {
                endRow++;
            }
            string sum = "=SUM({0}{1}:{0}{2})";
            Range footerEditRange = recapSheet.get_Range(string.Format(detailRowStart, endRow + 1), Type.Missing);
            foreach (var key in regularMap.Keys)
            {
                footerEditRange.Columns[startColumnDetail].Value = string.Format(sum, CommonUtil.GetExcelColumnName(startColumnDetail++), startRow, endRow);
                footerEditRange.Columns[startColumnDetail].Value = string.Format(sum, CommonUtil.GetExcelColumnName(startColumnDetail++), startRow, endRow);
                footerEditRange.Columns[startColumnDetail].Value = string.Format(sum, CommonUtil.GetExcelColumnName(startColumnDetail++), startRow, endRow);
            }
            foreach (var key in teamsMap.Keys)
            {
                footerEditRange.Columns[startColumnDetail].Value = string.Format(sum, CommonUtil.GetExcelColumnName(startColumnDetail++), startRow, endRow);
                footerEditRange.Columns[startColumnDetail].Value = string.Format(sum, CommonUtil.GetExcelColumnName(startColumnDetail++), startRow, endRow);
                footerEditRange.Columns[startColumnDetail].Value = string.Format(sum, CommonUtil.GetExcelColumnName(startColumnDetail++), startRow, endRow);
            }
            foreach (var key in offMap.Keys)
            {
                footerEditRange.Columns[startColumnDetail].Value = string.Format(sum, CommonUtil.GetExcelColumnName(startColumnDetail++), startRow, endRow);
                footerEditRange.Columns[startColumnDetail].Value = string.Format(sum, CommonUtil.GetExcelColumnName(startColumnDetail++), startRow, endRow);
                footerEditRange.Columns[startColumnDetail].Value = string.Format(sum, CommonUtil.GetExcelColumnName(startColumnDetail++), startRow, endRow);
            }
            // TOTAL
            footerEditRange.Columns[startColumnDetail].Value = string.Format(sum, CommonUtil.GetExcelColumnName(startColumnDetail++), startRow, endRow);
            footerEditRange.Columns[startColumnDetail].Value = string.Format(sum, CommonUtil.GetExcelColumnName(startColumnDetail++), startRow, endRow);
            footerEditRange.Columns[startColumnDetail].Value = string.Format(sum, CommonUtil.GetExcelColumnName(startColumnDetail++), startRow, endRow);
            //OFF
            footerEditRange.Columns[startColumnDetail].Value = string.Format(sum, CommonUtil.GetExcelColumnName(startColumnDetail++), startRow, endRow);
            //AL
            footerEditRange.Columns[startColumnDetail].Value = string.Format(sum, CommonUtil.GetExcelColumnName(startColumnDetail++), startRow, endRow);
            //SL
            footerEditRange.Columns[startColumnDetail].Value = string.Format(sum, CommonUtil.GetExcelColumnName(startColumnDetail++), startRow, endRow);
            //ML
            footerEditRange.Columns[startColumnDetail].Value = string.Format(sum, CommonUtil.GetExcelColumnName(startColumnDetail++), startRow, endRow);
            //P
            footerEditRange.Columns[startColumnDetail].Value = string.Format(sum, CommonUtil.GetExcelColumnName(startColumnDetail++), startRow, endRow);
            //WP
            footerEditRange.Columns[startColumnDetail].Value = string.Format(sum, CommonUtil.GetExcelColumnName(startColumnDetail++), startRow, endRow);

            #endregion
        }

        private string GetColumnName(Range range)
        {
            string name = range.Address;
            return name;
        }

        private void SetTimeTable(WorkDataModel workData)
        {
            string timeTableDsp = string.Empty;
            workData.IsLate = false;
            workData.IsLeaveEarly = false;
            if (workData.Time_table_no != null)
            {

                var timeTable = timeTableList.Find(x => x.Time_table_no == workData.Time_table_no);
                if (timeTable != null)
                {
                    timeTableDsp = CommonUtil.GetTimeShiftDsp(timeTable);
                    workData.Work_from = timeTable.Work_from;
                    workData.Work_to = timeTable.Work_to;
                    if (workData.WorkingType == WorkingType.Present)
                    {
                        decimal? update_start_time = CommonUtil.ToNullableDecimal(workData.Update_start_time);
                        if (update_start_time != null && timeTable.Work_from < update_start_time)
                        {
                            workData.Being_late_time = CommonUtil.SubTime(update_start_time.Value, timeTable.Work_from).ToString();
                            workData.IsLate = true;
                        }
                        decimal? leaving_early_time = null;
                        if (string.IsNullOrEmpty(workData.Leaving_early_time))
                        {
                            decimal? update_end_time = CommonUtil.ToNullableDecimal(workData.Update_end_time);
                            if (update_end_time != null && update_end_time < timeTable.Work_to)
                            {
                                leaving_early_time = CommonUtil.SubTime(timeTable.Work_to, update_end_time.Value);
                            }
                        }
                        else
                        {
                            leaving_early_time = CommonUtil.ToNullableDecimal(workData.Leaving_early_time);
                        }
                        if (leaving_early_time != null && leaving_early_time > 15)
                        {
                            workData.IsLeaveEarly = true;
                        }
                        else
                        {
                            workData.IsLeaveEarly = false;
                        }

                    }
                }
            }
            workData.TimeTable = timeTableDsp;
        }
        private string FormatHeaderByDay(string header)
        {
            if (string.IsNullOrEmpty(header))
            {
                return header;
            }
            string[] parts = header.Split(':');
            string format = "yyyy.MM.dd";
            if (parts.Length == 2)
            {
                format = parts[1].Trim();

            }
            return string.Format(parts[0], CommonUtil.ToDateTime(currentYearMonth).ToString(format)).ToUpper();

        }
        private string FormatHeaderByMonth(string header)
        {
            if (string.IsNullOrEmpty(header))
            {
                return header;
            }

            return string.Format(header, CommonUtil.ToDateTime(currentYearMonth).ToString("YM")).ToUpper();

        }
        private string FormatHeaderForMonthlyReport(string format, string companyNameOrPostName)
        {
            if (string.IsNullOrEmpty(format))
            {
                return string.Empty;
            }
            return string.Format(format, companyNameOrPostName, CommonUtil.GetLongYearMonth(selectedDay2Index + 1, selectedDay1)).ToUpper();
        }

        public void MonthlyExport()
        {
            Message = null;
            currentYearMonth = CommonUtil.GetDateAsDecimal(selectedDay1, CommonUtil.ToString(selectedDay2Index + 1), CommonUtil.ToString(selectedDay3Index + 1));


            workDataList.Sort(delegate(WorkDataModel o1, WorkDataModel o2)
            {
                string post_name1 = o1.Post_name == null ? "" : o1.Post_name;
                string post_name2 = o2.Post_name == null ? "" : o2.Post_name;
                int compare = post_name1.CompareTo(post_name2);
                if (compare == 0)
                {
                    compare = o1.Employee_no.CompareTo(o2.Employee_no);
                }
                return compare;
            });
            
            var checkedList = workDataList.FindAll(x => x.IsSelected).Select(x => x.Employee_no).ToList();
            if (checkedList.Count == 0)
            {
                Message = ResourcesManager.GetMessage(MessageConstant.A0011);
                return;
            }

            string templateFile = Properties.Settings.Default.XLS_Monthly_Report;
            if (Properties.Settings.Default.XLS_Use_Multi_Language)
            {
                templateFile = string.Format(templateFile, Properties.Settings.Default.SelectedLanguage);
            }
            else
            {
                templateFile = string.Format(templateFile, string.Empty);
            }
            string outFileName = selectedCompany.Company_name;
            bool isAllDepartment = false;
            if (selectedDepartment.Post_no != "0")
            {
                isAllDepartment = false;
                outFileName += " , " + selectedDepartment.Post_name;
            }
            else
            {
                isAllDepartment = true;
            }

            List<WorkDataModel> monthlyList = new List<WorkDataModel>();
            foreach (var employee in checkedList)
            {
                monthlyList.AddRange(workDataAllList.FindAll(x => x.Employee_no == employee));
            }

            templateFile = CommonUtil.GetTemplate(templateFile);
            outFileName = CommonUtil.Capitalize(outFileName);
            outFileName = string.Format(Properties.Settings.Default.XLS_Out_Monthly_File, outFileName, CommonUtil.GetDateAsString(currentYearMonth, "MY").ToUpper() + ".");

            try
            {
                CreateExcelApplication();
                workBook = CommonUtil.CreateWorkbook(excelApp, outFileName);

                templateWorkbook = excelApp.Workbooks.Open(templateFile, 0, false, 5, "", "", false, XlPlatform.xlWindows, "", true, false, 0, true, false, false);

                Worksheet tempMonthlyAttendanceSheet = null;
                Worksheet tempMonthlyViolationSheet = null;
                Worksheet monthlyAttendanceSheet;
                Worksheet monthlyViolationSheet;

                if (!isAllDepartment)
                {
                    tempMonthlyAttendanceSheet = templateWorkbook.Worksheets[DBConstant.TEMP_MONTHLY_ATTENDANCE_BY_DEPART];
                    tempMonthlyViolationSheet = templateWorkbook.Worksheets[DBConstant.TEMP_MONTHLY_VIOLATION_BY_DEPART];
                    tempMonthlyAttendanceSheet.Copy(workBook.Worksheets[Properties.Settings.Default.XLS_Default_Sheet_Name], Type.Missing);
                    tempMonthlyViolationSheet.Copy(workBook.Worksheets[Properties.Settings.Default.XLS_Default_Sheet_Name], Type.Missing);
                    monthlyAttendanceSheet = (Worksheet)workBook.Worksheets[tempMonthlyAttendanceSheet.Name];
                    monthlyViolationSheet = (Worksheet)workBook.Worksheets[tempMonthlyViolationSheet.Name];
                }
                else
                {
                    tempMonthlyAttendanceSheet = templateWorkbook.Worksheets[DBConstant.TEMP_MONTHLY_ATTENDANCE_BY_COMPANY];
                    tempMonthlyViolationSheet = templateWorkbook.Worksheets[DBConstant.TEMP_MONTHLY_VIOLATION_BY_COMPANY];
                    tempMonthlyAttendanceSheet.Copy(workBook.Worksheets[Properties.Settings.Default.XLS_Default_Sheet_Name], Type.Missing);
                    tempMonthlyViolationSheet.Copy(workBook.Worksheets[Properties.Settings.Default.XLS_Default_Sheet_Name], Type.Missing);
                    monthlyAttendanceSheet = (Worksheet)workBook.Worksheets[tempMonthlyAttendanceSheet.Name];
                    monthlyViolationSheet = (Worksheet)workBook.Worksheets[tempMonthlyViolationSheet.Name];
                }
                monthlyAttendanceSheet.Name = DBConstant.SHEET_MONTHLY_ATTENDANCE;
                monthlyViolationSheet.Name = DBConstant.SHEET_MONTHLY_VIOLATION;


                List<WorkDataModel> violations = monthlyList.FindAll(x => !x.IsOutOfExpiration && (
                     x.IsNoOnDuty || x.IsNoOffDuty || x.IsLate || x.IsLeaveEarly));
                List<string> violationsEmployeeCode = violations.GroupBy(x => x.Employee_no).Select(y => y.Key).ToList();
                CommonUtil.SetXlsPageSetup(monthlyAttendanceSheet, XlPageOrientation.xlLandscape);
                CommonUtil.SetXlsPageSetup(monthlyViolationSheet, XlPageOrientation.xlPortrait);
                monthlyAttendanceSheet.PageSetup.Application.MeasurementUnit = 1;
                monthlyAttendanceSheet.PageSetup.HeaderMargin = 0;
                monthlyAttendanceSheet.PageSetup.TopMargin = 85.04;
                SetMonthlyAttendanceSheet(monthlyAttendanceSheet, isAllDepartment, checkedList);
                SetMonthlyViolationsSheet(monthlyViolationSheet, isAllDepartment, violations, violationsEmployeeCode);
                CommonUtil.DeleteDefaultSheet(workBook);
                ((Microsoft.Office.Interop.Excel._Worksheet)monthlyAttendanceSheet).Activate();
                this.templateWorkbook.Saved = true;
                this.templateWorkbook.Close();
                this.Show();

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                if (workBook != null)
                {
                    workBook.Saved = true;
                    workBook.Close();
                }
            }
            finally
            {
                DisposeXls();
            }
        }

        private void SetMonthlyAttendanceSheet(Worksheet monthlySheet, bool isAllDepartment, List<string> employeeNoList)
        {

            Range headerNameRange = monthlySheet.get_Range("A3", Type.Missing);
            Range detailRange = monthlySheet.get_Range("A6", Type.Missing);
            Range detailTemRange = monthlySheet.get_Range("A6", Type.Missing);
            monthlySheet.PageSetup.Zoom = false;
            monthlySheet.PageSetup.FitToPagesWide = 1;
            monthlySheet.PageSetup.FitToPagesTall = 100;

            monthlySheet.PageSetup.Orientation = XlPageOrientation.xlLandscape;
            monthlySheet.PageSetup.PaperSize = XlPaperSize.xlPaperA4;
            monthlySheet.PageSetup.RightMargin = 10;
            monthlySheet.PageSetup.LeftMargin = 20;

            if (isAllDepartment)
            {
                headerNameRange.Columns[1].Value = FormatHeaderForMonthlyReport(headerNameRange.Columns[1].Value, selectedCompany.Company_name);
            }
            else
            {
                headerNameRange.Columns[1].Value = FormatHeaderForMonthlyReport(headerNameRange.Columns[1].Value, selectedDepartment.Post_name);
            }
            decimal fristDayOfMonth = CommonUtil.GetFirstDayOfMonth(currentYearMonth);
            decimal lastDayOfMonth = CommonUtil.GetLastDayOfMonth(currentYearMonth);
            int startRow = 6;
            int rowIndex = startRow;
            int headerRowIndex = 5;
            string detailRowStart = "A{0}";
            int count = 0;
            int totalRow = employeeNoList.Count + rowIndex; ;

            #region create header
            int startHeaderDayColumn = 5;
            if (!isAllDepartment)
            {
                startHeaderDayColumn = 4;
            }
            int dayCount = 0;

            for (decimal day = fristDayOfMonth; day <= lastDayOfMonth; day++)
            {
                dayCount++;

                var holiday = holidayList.Find(x => x.Holiday_date == day);
                if (holiday != null)
                {
                    if (holiday.National_holiday_flag == DBConstant.HOLIDAY_FLAG_NATIONAL)
                    {
                        monthlySheet.Cells[headerRowIndex, startHeaderDayColumn + dayCount].Interior.Color = CommonUtil.GetExcelColor(DBConstant.COLOR_HOLIDAY_NATIONAL);
                        monthlySheet.Cells[headerRowIndex + 1, startHeaderDayColumn + dayCount].Interior.Color = CommonUtil.GetExcelColor(DBConstant.COLOR_HOLIDAY_NATIONAL);
                    }
                    else
                    {
                        monthlySheet.Cells[headerRowIndex, startHeaderDayColumn + dayCount].Interior.Color = CommonUtil.GetExcelColor(DBConstant.COLOR_HOLIDAY_WEEKEND);
                        monthlySheet.Cells[headerRowIndex + 1, startHeaderDayColumn + dayCount].Interior.Color = CommonUtil.GetExcelColor(DBConstant.COLOR_HOLIDAY_WEEKEND);
                    }
                }
            }
            if (dayCount < 31)
            {
                int removeColumn = dayCount + 1;
                for (int i = removeColumn; i <= 31; i++)
                {
                    string columnName = CommonUtil.GetExcelColumnName(removeColumn + startHeaderDayColumn);
                    monthlySheet.Range[columnName + 1].EntireColumn.Delete();
                }
            }
            #endregion
            CommonUtil.CreateEmptyRow(monthlySheet, startRow, employeeNoList.Count);
            foreach (var employeeNo in employeeNoList)
            {
                EmployeeModel employee = employeeList.Find(x => x.Employee_no == employeeNo);
                int columnIndex = 1;

                Range detailEditRange = monthlySheet.get_Range(string.Format(detailRowStart, rowIndex), Type.Missing);
                detailEditRange.Columns[columnIndex++].Value = ++count;
                if (isAllDepartment)
                {
                    detailEditRange.Columns[columnIndex++].Value = CommonUtil.GetXlsValue(employee.Post_name);
                }
                detailEditRange.Columns[columnIndex++].Value = CommonUtil.GetFullName(employee.Emsize_first_name, employee.Emsize_last_name);
                detailEditRange.Columns[columnIndex++].Value = employee.Employee_no;
                detailEditRange.Columns[columnIndex++].Value = CommonUtil.GetPosition(employee.Remarks);

                int totalAL = 0;
                int totalSL = 0;
                int totalML = 0;
                decimal totalP = 0;
                int totalWP = 0;
                decimal totalD = 0;
                decimal totalDH = 0;
                decimal totalDNH = 0;
                decimal totalN = 0;
                decimal totalNH = 0;
                decimal totalNNH = 0;
                decimal totalOTD = 0;
                decimal totalOTN = 0;
                decimal totalOTH = 0;
                decimal totalOTNH = 0;
                decimal totalOTNIH = 0;
                decimal totalOTNINH = 0;
                decimal totalNormalWorking = 0;
                for (decimal day = fristDayOfMonth; day <= lastDayOfMonth; day++)
                {
                    var workData = workDataAllList.Find(x => x.Work_date == day && x.Employee_no == employee.Employee_no && !x.IsOutOfExpiration);
                    if (workData != null && workData.Time_table_no != null)
                    {
                        TimeTableModel timeTable = timeTableList.Find(x => x.Time_table_no == workData.Time_table_no);
                        if (timeTable == null)
                        {
                            continue;
                        }

                        decimal hourTime; // minutes
                        decimal? contact_time = CommonUtil.ToNullableDecimal(workData.Contract_time);
                        if (contact_time == null || contact_time == 0)
                        {
                            hourTime = CommonUtil.GetContactTime(timeTable);
                        }
                        else
                        {
                            hourTime = contact_time.Value;
                        }
                        if (hourTime == 0)
                        {
                            hourTime = 480;
                        }
                        string dayWork = "";
                        switch ((int)workData.Work_type_no)
                        {
                            case DBConstant.WORK_TYPE_NORMAL:
                                if (workData.Working_time != null && !CommonUtil.ToDecimal(workData.Working_time).Equals(0))
                                {
                                    decimal result = Math.Round(CommonUtil.ToDecimal(workData.Working_time) / hourTime, 2);
                                    string dayName = "D";
                                    HolidayModel holiday = holidayList.Find(x => x.Holiday_date == workData.Work_date);
                                    if (holiday != null && employee.Use_flag_of_holiday == 1 || workData.Work_day_type_no > DBConstant.WORK_DAY_TYPE_NORMAL)
                                    {
                                        if ((holiday != null && holiday.National_holiday_flag == DBConstant.HOLIDAY_FLAG_NATIONAL) || workData.Work_day_type_no == 7)
                                        {
                                            // work at night on national holiday
                                            if (CommonUtil.IsMiddleNightShift(timeTable))
                                            {
                                                dayName = "NNH";
                                                totalNNH += result;
                                                totalOTNINH += CommonUtil.GetOverTime(result);
                                            }
                                            else
                                            {
                                                // work on national holiday
                                                dayName = "DNH";
                                                totalDNH += result;
                                                totalOTNH += CommonUtil.GetOverTime(result);
                                            }

                                        }
                                        else
                                        {
                                            if (CommonUtil.IsMiddleNightShift(timeTable))
                                            {
                                                dayName = "NH";
                                                totalNH += result;
                                                totalOTNIH += CommonUtil.GetOverTime(result);
                                            }
                                            else
                                            {
                                                totalDH += result;
                                                totalOTH += CommonUtil.GetOverTime(result);
                                                dayName = "DH";
                                            }


                                        }

                                    }
                                    else
                                    {
                                        if (CommonUtil.IsMiddleNightShift(timeTable))
                                        {
                                            dayName = "N";
                                            totalN += result;
                                            totalOTN += CommonUtil.GetOverTime(result);
                                        }
                                        else
                                        {
                                            totalD += result;
                                            totalOTD += CommonUtil.GetOverTime(result);
                                            totalNormalWorking++;
                                        }

                                    }
                                    if (result != 1)
                                    {
                                        dayWork = result + dayName;
                                    }
                                    else
                                    {
                                        dayWork = dayName;
                                    }

                                }
                                else if (workData.WorkingType == WorkingType.Unknown)
                                {
                                    if (workData.IsOutOfExpiration)
                                    {
                                        dayWork = string.Empty;
                                    }
                                    else
                                    {
                                        dayWork = "-";
                                        //totalNormalWorking++;
                                    }
                                }

                                break;
                            case DBConstant.WORK_TYPE_ANNUAL_LEAVE:
                                dayWork = DBConstant.ABSENT_ANNUAL_LEAVE;
                                totalAL++;
                                break;
                            case DBConstant.WORK_TYPE_SPECIAL_LEAVE:
                                dayWork = DBConstant.ABSENT_SPECIAL_LEAVE;
                                totalSL++;
                                break;
                            case DBConstant.WORK_TYPE_MATERNITY_LEAVE:
                                dayWork = DBConstant.ABSENT_MATERNITY_LEAVE;
                                totalML++;
                                break;
                            case DBConstant.WORK_TYPE_PERMISSION:
                                dayWork = DBConstant.ABSENT_PERMISSION;
                                totalP++;
                                break;
                            case DBConstant.WORK_TYPE_AW_PERMISSION:
                                dayWork = DBConstant.ABSENT_WITHOUT_PERMISSION;
                                totalWP++;
                                break;
                            case DBConstant.WORK_TYPE_HALF_PERMISSION:
                                dayWork = DBConstant.WORKING_HALF_PERMISION + "D";
                                totalP += DBConstant.WORKING_HALF_PERMISION;
                                totalD += DBConstant.WORKING_HALF_PERMISION;
                                totalNormalWorking += DBConstant.WORKING_HALF_PERMISION;
                                break;
                            case DBConstant.WORK_TYPE_HOLIDAY_DUTY:
                                if (workData.Holiday_time != null)
                                {
                                    decimal result = Math.Round(CommonUtil.ToDecimal(workData.Holiday_time) / hourTime, 2);
                                    string dayName = "DH";
                                    HolidayModel holiday = holidayList.Find(x => x.Holiday_date == workData.Work_date);
                                    if (holiday != null)
                                    {
                                        if (holiday.National_holiday_flag == DBConstant.HOLIDAY_FLAG_NATIONAL)
                                        {
                                            if (CommonUtil.IsMiddleNightShift(timeTable))
                                            {
                                                dayName = "NNH";
                                                totalNNH += result;
                                                totalOTNINH += CommonUtil.GetOverTime(result);
                                            }
                                            else
                                            {
                                                dayName = "DNH";
                                                totalDNH += result;
                                                totalOTH += CommonUtil.GetOverTime(result);
                                            }

                                        }
                                        else
                                        {
                                            if (CommonUtil.IsMiddleNightShift(timeTable))
                                            {
                                                dayName = "NH";
                                                totalNH += result;
                                                totalOTNIH += CommonUtil.GetOverTime(result);
                                            }
                                            else
                                            {
                                                totalDH += result;
                                                totalOTH += CommonUtil.GetOverTime(result);
                                            }

                                        }
                                    }
                                    else
                                    {
                                        if (workData.Work_day_type_no == 7)
                                        {
                                            if (CommonUtil.IsMiddleNightShift(timeTable))
                                            {
                                                dayName = "NNH";
                                                totalNNH += result;
                                                totalOTNINH += CommonUtil.GetOverTime(result);
                                            }
                                            else
                                            {
                                                dayName = "DNH";
                                                totalDNH += result;
                                                totalOTH += CommonUtil.GetOverTime(result);
                                            }

                                        }
                                        else
                                        {
                                            if (CommonUtil.IsMiddleNightShift(timeTable))
                                            {
                                                dayName = "NH";
                                                totalNH += result;
                                                totalOTNIH += CommonUtil.GetOverTime(result);
                                            }
                                            else
                                            {
                                                totalDH += result;
                                                totalOTH += CommonUtil.GetOverTime(result);
                                            }

                                        }
                                    }

                                    if (result != 1)
                                    {
                                        dayWork = result + dayName;
                                    }
                                    else
                                    {
                                        dayWork = dayName;
                                    }

                                }

                                break;
                        }

                        detailEditRange.Columns[columnIndex].Value = dayWork;
                    }
                    else
                    {
                        detailEditRange.Columns[columnIndex].Value = string.Empty;
                    }
                    columnIndex++;

                }
                //total
                if (totalD > totalNormalWorking)
                {
                    totalOTD = totalD - totalNormalWorking;
                    totalD = totalNormalWorking;
                }

                // totalD = totalD - totalOTD;
                totalDH = totalDH - totalOTH;
                totalDNH = totalDNH - totalOTNH;
                totalN = totalN - totalOTN;

                totalOTNH = totalOTNH + totalOTNINH;
                totalOTH = totalOTH + totalOTNIH;

                detailEditRange.Columns[columnIndex++].Value = CommonUtil.GetXlsValue(totalD);
                detailEditRange.Columns[columnIndex++].Value = CommonUtil.GetXlsValue(totalDH);
                detailEditRange.Columns[columnIndex++].Value = CommonUtil.GetXlsValue(totalDNH);
                detailEditRange.Columns[columnIndex++].Value = CommonUtil.GetXlsValue(totalN);
                detailEditRange.Columns[columnIndex++].Value = CommonUtil.GetXlsValue(totalNH);
                detailEditRange.Columns[columnIndex++].Value = CommonUtil.GetXlsValue(totalNNH);
                detailEditRange.Columns[columnIndex++].Value = CommonUtil.GetXlsValue(totalAL);
                detailEditRange.Columns[columnIndex++].Value = CommonUtil.GetXlsValue(totalSL);
                detailEditRange.Columns[columnIndex++].Value = CommonUtil.GetXlsValue(totalML);
                detailEditRange.Columns[columnIndex++].Value = CommonUtil.GetXlsValue(totalP);
                detailEditRange.Columns[columnIndex++].Value = CommonUtil.GetXlsValue(totalWP);
                detailEditRange.Columns[columnIndex++].Value = CommonUtil.GetXlsValue(totalOTD);
                detailEditRange.Columns[columnIndex++].Value = CommonUtil.GetXlsValue(totalOTN);
                detailEditRange.Columns[columnIndex++].Value = CommonUtil.GetXlsValue(totalOTH);
                detailEditRange.Columns[columnIndex++].Value = CommonUtil.GetXlsValue(totalOTNH);
                rowIndex++;
            }


            string totalFormular = "=SUM({0}{1}:{0}{2})";
            int endRow = rowIndex - 1;
            int startTotalColumn = startHeaderDayColumn + dayCount + 1;
            monthlySheet.Cells[rowIndex, startTotalColumn].Value = string.Format(totalFormular, CommonUtil.GetExcelColumnName(startTotalColumn++), startRow, endRow);
            monthlySheet.Cells[rowIndex, startTotalColumn].Value = string.Format(totalFormular, CommonUtil.GetExcelColumnName(startTotalColumn++), startRow, endRow);
            monthlySheet.Cells[rowIndex, startTotalColumn].Value = string.Format(totalFormular, CommonUtil.GetExcelColumnName(startTotalColumn++), startRow, endRow);
            monthlySheet.Cells[rowIndex, startTotalColumn].Value = string.Format(totalFormular, CommonUtil.GetExcelColumnName(startTotalColumn++), startRow, endRow);
            monthlySheet.Cells[rowIndex, startTotalColumn].Value = string.Format(totalFormular, CommonUtil.GetExcelColumnName(startTotalColumn++), startRow, endRow);

            monthlySheet.Cells[rowIndex, startTotalColumn].Value = string.Format(totalFormular, CommonUtil.GetExcelColumnName(startTotalColumn++), startRow, endRow);
            monthlySheet.Cells[rowIndex, startTotalColumn].Value = string.Format(totalFormular, CommonUtil.GetExcelColumnName(startTotalColumn++), startRow, endRow);
            monthlySheet.Cells[rowIndex, startTotalColumn].Value = string.Format(totalFormular, CommonUtil.GetExcelColumnName(startTotalColumn++), startRow, endRow);
            monthlySheet.Cells[rowIndex, startTotalColumn].Value = string.Format(totalFormular, CommonUtil.GetExcelColumnName(startTotalColumn++), startRow, endRow);
            monthlySheet.Cells[rowIndex, startTotalColumn].Value = string.Format(totalFormular, CommonUtil.GetExcelColumnName(startTotalColumn++), startRow, endRow);

            monthlySheet.Cells[rowIndex, startTotalColumn].Value = string.Format(totalFormular, CommonUtil.GetExcelColumnName(startTotalColumn++), startRow, endRow);
            monthlySheet.Cells[rowIndex, startTotalColumn].Value = string.Format(totalFormular, CommonUtil.GetExcelColumnName(startTotalColumn++), startRow, endRow);
            monthlySheet.Cells[rowIndex, startTotalColumn].Value = string.Format(totalFormular, CommonUtil.GetExcelColumnName(startTotalColumn++), startRow, endRow);
            monthlySheet.Cells[rowIndex, startTotalColumn].Value = string.Format(totalFormular, CommonUtil.GetExcelColumnName(startTotalColumn++), startRow, endRow);
            monthlySheet.Cells[rowIndex, startTotalColumn].Value = string.Format(totalFormular, CommonUtil.GetExcelColumnName(startTotalColumn++), startRow, endRow);

        }
        private void SetMonthlyViolationsSheet(Worksheet monthlySheet, bool isAllDepartment, List<WorkDataModel> violations, List<string> employeeNoList)
        {

            Range headerNameRange = monthlySheet.get_Range("A2", Type.Missing);
            Range detailRange = monthlySheet.get_Range("A6", Type.Missing);
            Range detailTemRange = monthlySheet.get_Range("A6", Type.Missing);

            if (isAllDepartment)
            {
                headerNameRange.Columns[1].Value = FormatHeaderForMonthlyReport(headerNameRange.Columns[1].Value, selectedCompany.Company_name);
            }
            else
            {
                headerNameRange.Columns[1].Value = FormatHeaderForMonthlyReport(headerNameRange.Columns[1].Value, selectedDepartment.Post_name);
            }

            decimal fristDayOfMonth = CommonUtil.GetFirstDayOfMonth(currentYearMonth);
            decimal lastDayOfMonth = CommonUtil.GetLastDayOfMonth(currentYearMonth);
            int startRow = 5;
            int rowIndex = startRow;

            string detailRowStart = "A{0}";

            int count = 0;
            int totalRow = employeeNoList.Count + rowIndex; ;
            int startTotalColumn;
            CommonUtil.CreateEmptyRow(monthlySheet, startRow, employeeNoList.Count);
            foreach (var employeeNo in employeeNoList)
            {
                EmployeeModel employee = employeeList.Find(x => x.Employee_no == employeeNo);
                int totalNoOnDuty = violations.Count(x => x.Employee_no == employeeNo && x.IsNoOnDuty);
                int totalNoOffDuty = violations.Count(x => x.Employee_no == employeeNo && x.IsNoOffDuty);
                int totalLate = violations.Count(x => x.Employee_no == employeeNo && x.WorkingType == WorkingType.Present && x.IsLate);
                int totalLeftEarly = violations.Count(x => x.Employee_no == employeeNo && x.WorkingType == WorkingType.Present && x.IsLeaveEarly);

                int columnIndex = 1;
                Range detailEditRange = monthlySheet.get_Range(string.Format(detailRowStart, rowIndex), Type.Missing);

                detailEditRange.Columns[columnIndex++].Value = ++count;
                if (isAllDepartment)
                {
                    detailEditRange.Columns[columnIndex++].Value = CommonUtil.GetXlsValue(employee.Post_name);
                }
                detailEditRange.Columns[columnIndex++].Value = CommonUtil.GetFullName(employee.Emsize_first_name, employee.Emsize_last_name);
                detailEditRange.Columns[columnIndex++].Value = employee.Employee_no;
                detailEditRange.Columns[columnIndex++].Value = CommonUtil.GetPosition(employee.Remarks);

                detailEditRange.Columns[columnIndex++].Value = CommonUtil.GetXlsValue(totalNoOnDuty);
                detailEditRange.Columns[columnIndex++].Value = CommonUtil.GetXlsValue(totalLate);
                detailEditRange.Columns[columnIndex++].Value = CommonUtil.GetXlsValue(totalNoOffDuty);
                detailEditRange.Columns[columnIndex++].Value = CommonUtil.GetXlsValue(totalNoOnDuty + totalLate + totalNoOffDuty);
                detailEditRange.Columns[columnIndex++].Value = CommonUtil.GetXlsValue(totalLeftEarly);
                rowIndex++;
            }
            if (isAllDepartment)
            {
                startTotalColumn = 6;
            }
            else
            {
                startTotalColumn = 5;
            }
            string totalFormular = "=SUM({0}{1}:{0}{2})";
            int endRow = rowIndex - 1;
            if (totalRow == startRow)
            {
                totalRow++;
            }
            monthlySheet.Cells[totalRow, startTotalColumn].Value = string.Format(totalFormular, CommonUtil.GetExcelColumnName(startTotalColumn++), startRow, endRow);
            monthlySheet.Cells[totalRow, startTotalColumn].Value = string.Format(totalFormular, CommonUtil.GetExcelColumnName(startTotalColumn++), startRow, endRow);
            monthlySheet.Cells[totalRow, startTotalColumn].Value = string.Format(totalFormular, CommonUtil.GetExcelColumnName(startTotalColumn++), startRow, endRow);
            monthlySheet.Cells[totalRow, startTotalColumn].Value = string.Format(totalFormular, CommonUtil.GetExcelColumnName(startTotalColumn++), startRow, endRow);
            monthlySheet.Cells[totalRow, startTotalColumn].Value = string.Format(totalFormular, CommonUtil.GetExcelColumnName(startTotalColumn++), startRow, endRow);
        }

        protected void Dispose()
        {

            if (excelApp != null)
            {
                Marshal.ReleaseComObject(excelApp);
            }
        }

        private void CreateExcelApplication()
        {

            excelApp = (Microsoft.Office.Interop.Excel.Application)RuntimeHelpers.GetObjectValue(new Microsoft.Office.Interop.Excel.Application());
            excelApp.DisplayAlerts = false;
            excelApp.Visible = false;
        }

        private void Show()
        {
            try
            {
                this.excelApp.DisplayAlerts = true;
                this.excelApp.WindowState = XlWindowState.xlMaximized;
                this.excelApp.Visible = true;
                //this.excelApp.SendKeys("{ESC}");
                foreach (Process process in Process.GetProcessesByName("Excel"))
                {
                    if (process.MainWindowTitle.TrimEnd(new char[0]).EndsWith(((Workbook)this.workBook).FullName))
                    {
                        CommonUtil.SetWindowPosision(0L, process.Handle);
                        return;
                    }
                }
            }
            finally
            {

            }

        }

        public void Cancel()
        {
            this.TryClose(true);
        }
        #region get/set
        public bool IsAllChecked
        {
            get
            {
                return isAllChecked;
            }
            set
            {
                if (isAllChecked != value)
                {
                    isAllChecked = value;
                    NotifyOfPropertyChange(() => IsAllChecked);
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

        public LanguageModel LblEmployeeList
        {
            get
            {
                return lblEmployeeList;
            }
            set
            {
                if (lblEmployeeList != value)
                {
                    lblEmployeeList = value;
                    NotifyOfPropertyChange(() => LblEmployeeList);
                }
            }
        }

        public LanguageModel LblCompany
        {
            get
            {
                return lblCompany;
            }
            set
            {
                if (lblCompany != value)
                {
                    lblCompany = value;
                    NotifyOfPropertyChange(() => LblCompany);
                }
            }
        }

        public LanguageModel LblDepartment
        {
            get
            {
                return lblDepartment;
            }
            set
            {
                if (lblDepartment != value)
                {
                    lblDepartment = value;
                    NotifyOfPropertyChange(() => LblDepartment);
                }
            }
        }

        public LanguageModel LblEmployeeSearch
        {
            get
            {
                return lblEmployeeSearch;
            }
            set
            {
                if (lblEmployeeSearch != value)
                {
                    lblEmployeeSearch = value;
                    NotifyOfPropertyChange(() => LblEmployeeSearch);
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

        public LanguageModel LblYearMonth
        {
            get
            {
                return lblYearMonth;
            }
            set
            {
                if (lblYearMonth != value)
                {
                    lblYearMonth = value;
                    NotifyOfPropertyChange(() => LblYearMonth);
                }
            }
        }

        public LanguageModel LblSearch
        {
            get
            {
                return lblSearch;
            }
            set
            {
                if (lblSearch != value)
                {
                    lblSearch = value;
                    NotifyOfPropertyChange(() => LblSearch);
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

        public LanguageModel LblDailyExport
        {
            get
            {
                return lblDailyExport;
            }
            set
            {
                if (lblDailyExport != value)
                {
                    lblDailyExport = value;
                    NotifyOfPropertyChange(() => LblDailyExport);
                }
            }
        }

        public LanguageModel LblMonthlyExport
        {
            get
            {
                return lblMonthlyExport;
            }
            set
            {
                if (lblMonthlyExport != value)
                {
                    lblMonthlyExport = value;
                    NotifyOfPropertyChange(() => LblMonthlyExport);
                }
            }
        }

        public LanguageModel LblLogout
        {
            get
            {
                return lblLogout;
            }
            set
            {
                if (lblLogout != value)
                {
                    lblLogout = value;
                    NotifyOfPropertyChange(() => LblLogout);
                }
            }
        }

        public LanguageModel LblClose
        {
            get
            {
                return lblClose;
            }
            set
            {
                if (lblClose != value)
                {
                    lblClose = value;
                    NotifyOfPropertyChange(() => LblClose);
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
        public LanguageModel LblCheckAll
        {
            get
            {
                return lblCheckAll;
            }
            set
            {
                if (lblCheckAll != value)
                {
                    lblCheckAll = value;
                    NotifyOfPropertyChange(() => LblCheckAll);
                }
            }
        }
        public List<CompanyModel> CompanyList
        {
            get
            {
                return companyList;
            }
            set
            {
                if (companyList != value)
                {
                    companyList = value;
                    NotifyOfPropertyChange(() => CompanyList);
                }
            }
        }

        public List<PostModel> DepartmentList
        {
            get
            {
                return departmentList;
            }
            set
            {
                if (departmentList != value)
                {
                    departmentList = value;
                    NotifyOfPropertyChange(() => DepartmentList);
                }
            }
        }

        public CompanyModel SelectedCompany
        {
            get
            {
                return selectedCompany;
            }
            set
            {
                if (selectedCompany != value)
                {
                    selectedCompany = value;
                    NotifyOfPropertyChange(() => SelectedCompany);
                }
            }
        }

        public PostModel SelectedDepartment
        {
            get
            {
                return selectedDepartment;
            }
            set
            {
                if (selectedDepartment != value)
                {
                    selectedDepartment = value;
                    NotifyOfPropertyChange(() => SelectedDepartment);
                }
            }
        }
        public List<string> Day1List
        {
            get
            {
                return day1List;
            }
            set
            {
                if (day1List != value)
                {
                    day1List = value;
                    NotifyOfPropertyChange(() => Day1List);
                }
            }
        }

        public List<string> Day2List
        {
            get
            {
                return day2List;
            }
            set
            {
                if (day2List != value)
                {
                    day2List = value;
                    NotifyOfPropertyChange(() => Day2List);
                }
            }
        }

        public List<string> Day3List
        {
            get
            {
                return day3List;
            }
            set
            {
                if (day3List != value)
                {
                    day3List = value;
                    NotifyOfPropertyChange(() => Day3List);
                }
            }
        }

        public string SelectedDay1
        {
            get
            {
                return selectedDay1;
            }
            set
            {
                if (selectedDay1 != value)
                {
                    selectedDay1 = value;
                    NotifyOfPropertyChange(() => SelectedDay1);
                }
            }
        }

        public string SelectedDay2
        {
            get
            {
                return selectedDay2;
            }
            set
            {
                if (selectedDay2 != value)
                {
                    selectedDay2 = value;
                    NotifyOfPropertyChange(() => SelectedDay2);
                }
            }
        }

        public string SelectedDay3
        {
            get
            {
                return selectedDay3;
            }
            set
            {
                if (selectedDay3 != value)
                {
                    selectedDay3 = value;
                    NotifyOfPropertyChange(() => SelectedDay3);
                }
            }
        }


        public string EmployeeSearch
        {
            get
            {
                return employeeSearch;
            }
            set
            {
                if (employeeSearch != value)
                {
                    employeeSearch = value;
                    NotifyOfPropertyChange(() => EmployeeSearch);
                }
            }
        }

        public int SelectedDay1Index
        {
            get
            {
                return selectedDay1Index;
            }
            set
            {
                if (selectedDay1Index != value)
                {
                    selectedDay1Index = value;
                    NotifyOfPropertyChange(() => SelectedDay1Index);
                }
            }
        }

        public int SelectedDay2Index
        {
            get
            {
                return selectedDay2Index;
            }
            set
            {
                if (selectedDay2Index != value)
                {
                    selectedDay2Index = value;
                    NotifyOfPropertyChange(() => SelectedDay2Index);
                }
            }
        }

        public int SelectedDay3Index
        {
            get
            {
                return selectedDay3Index;
            }
            set
            {
                if (selectedDay3Index != value)
                {
                    selectedDay3Index = value;
                    NotifyOfPropertyChange(() => SelectedDay3Index);
                }
            }
        }
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
        public int ColYearListIndex
        {
            get
            {
                return colYearListIndex;
            }
            set
            {
                if (colYearListIndex != value)
                {
                    colYearListIndex = value;
                    NotifyOfPropertyChange(() => ColYearListIndex);
                }
            }
        }

        public int ColMonthListIndex
        {
            get
            {
                return colMonthListIndex;
            }
            set
            {
                if (colMonthListIndex != value)
                {
                    colMonthListIndex = value;
                    NotifyOfPropertyChange(() => ColMonthListIndex);
                }
            }
        }

        public int ColDayListIndex
        {
            get
            {
                return colDayListIndex;
            }
            set
            {
                if (colDayListIndex != value)
                {
                    colDayListIndex = value;
                    NotifyOfPropertyChange(() => ColDayListIndex);
                }
            }
        }
        #endregion
    }
}
