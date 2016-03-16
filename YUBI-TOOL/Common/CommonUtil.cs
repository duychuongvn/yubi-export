using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Windows.Data;
using Microsoft.Office.Interop.Excel;
using YUBI_TOOL.Model;

namespace YUBI_TOOL.Common
{
    public class CommonUtil
    {
        public static List<SelectItemModel> CreateSelectItemList(string displayText)
        {
            List<SelectItemModel> list = new List<SelectItemModel>();
            string[] text = displayText.Split(',');
            foreach (var value in text)
            {
                string[] parts = value.Trim().Split('=');
                SelectItemModel dayType = new SelectItemModel()
                {
                    ItemCD = parts[1].Trim(),
                    ItemValue = parts[0].Trim()
                };
                list.Add(dayType);
            }
            return list;
        }
        public static object GetPropertyValue(object obj, string propName)
        {
            if (obj == null)
            {
                return null;
            }
            var prop = obj.GetType().GetProperty(propName);
            return prop.GetValue(obj, null);
        }

        public static string GetCaption(string caption)
        {
            if (!string.IsNullOrEmpty(caption))
            {
                return Regex.Replace(caption, "[:：]$", "");
            }
            return caption;
        }
        public static string ToString(object value)
        {
            if (value != null)
            {
                return value.ToString();
            }
            return null;
        }

        public static decimal MinuteToHr(decimal minute)
        {
            decimal hr = decimal.Floor(minute / 60);
            hr = hr * 100 + minute - hr * 60;
            return hr;
        }
        public static decimal HourToMinute(decimal hour)
        {
            decimal minute = decimal.Floor(hour / 100);
            minute = minute * 60 + (hour - minute * 100);
            return minute;
        }


        public static decimal GetLastMonth(decimal currentMonth)
        {
            var date = CommonUtil.ToDateTime(currentMonth);
            return ToDecimal(date.AddMonths(-1));
        }

        public static decimal GetNextMonth(decimal currentMonth)
        {
            var date = CommonUtil.ToDateTime(currentMonth);
            return ToDecimal(date.AddMonths(1));
        }
        public static decimal GetFirstDayOfMonth(decimal currentMonth)
        {
            var date = CommonUtil.ToDateTime(currentMonth);
            var firstdate = new DateTime(date.Year, date.Month, 1);
            return ToDecimal(firstdate);
        }
        public static decimal GetLastDayOfMonth(decimal currentMonth)
        {
            var date = CommonUtil.ToDateTime(currentMonth);
            var firstdate = new DateTime(date.Year, date.Month, 1);
            var lastdate = firstdate.AddMonths(1).AddDays(-1);
            return ToDecimal(lastdate);
        }

        public static string GetDateAsStringWithFormat(decimal date, string format)
        {
            DateTime dateTime = ToDateTime(date);
            if (string.IsNullOrEmpty(format))
            {
                format = "yyyy.MM.dd";
            }
            return dateTime.ToString(format);
        }

        public static string GetLongMonthString(int month)
        {
            return ResourcesManager.MONTH_LONG[month - 1];
        }

        public static string GetLongYearMonth(int month, string year)
        {
            return GetLongMonthString(month).ToUpper() + " " + year;
        }

        public static decimal GetContactTime(TimeTableModel timeTable)
        {
            decimal hr = SubTime(timeTable.Work_to, timeTable.Work_from);
            string restFrom = "Rest{0}_from";
            string restTo = "Rest{0}_to";
            for (int i = 1; i <= 10; i++)
            {
                decimal? restTimeFrom = (decimal?)CommonUtil.GetPropertyValue(timeTable, string.Format(restFrom, i));
                decimal? restTimeTo = (decimal?)CommonUtil.GetPropertyValue(timeTable, string.Format(restTo, i));

                if (restTimeFrom != null && restTimeTo != null
                    && restTimeFrom.Value > timeTable.Work_from && restTimeFrom.Value < timeTable.Work_to
                    && restTimeTo.Value > timeTable.Work_from && restTimeTo.Value < timeTable.Work_to
                    )
                {
                    hr = hr - SubTime(restTimeTo.Value, restTimeFrom.Value);
                }
            }
            return hr;
        }
        public static string GetDateAsString(decimal date, string format, bool hasWeekDay = false)
        {
            string dateValue = null;
            if (string.IsNullOrEmpty(format))
            {
                return null;
            }
            format = format.ToUpper();
            DateTime dateTime = ToDateTime(date);
            Match dayMonthYear = Regex.Match(format, "^[D]{1,2}[M]{1,2}[Y]{1,4}$");
            Match yearMonthDay = Regex.Match(format, "^[Y]{1,4}[M]{1,2}[D]{1,2}$");
            Match dayMonth = Regex.Match(format, "^[D]{1,2}[M]{1,2}$");
            Match monthDay = Regex.Match(format, "^[M]{1,2}[D]{1,2}$");
            Match yearMonth = Regex.Match(format, "^[Y]{1,4}[Y]{1,2}$");
            Match monthYear = Regex.Match(format, "^[M]{1,2}[Y]{1,4}$");
            if (dayMonth.Success)
            {

                dateValue = dateTime.ToString("dd");
                dateValue += "/" + ResourcesManager.MONTH_SHORT[dateTime.Month - 1];
                if (hasWeekDay)
                {
                    dateValue += string.Format("({0})", ResourcesManager.WEEK_DAY_SHORT[(int)dateTime.DayOfWeek]);
                }
            }
            else if (monthDay.Success)
            {
                dateValue = dateTime.ToString("MM/dd");
                if (hasWeekDay)
                {
                    dateValue += string.Format("({0})", ResourcesManager.WEEK_DAY_SHORT[(int)dateTime.DayOfWeek]);
                }
            }
            else if (dayMonthYear.Success)
            {
                dateValue = dateTime.ToString("dd");
                dateValue += "/" + ResourcesManager.MONTH_SHORT[dateTime.Month - 1];
                dateValue += "/" + dateTime.Year;
                if (hasWeekDay)
                {
                    dateValue += string.Format("({0})", ResourcesManager.WEEK_DAY_SHORT[(int)dateTime.DayOfWeek]);
                }
            }
            else if (yearMonthDay.Success)
            {
                dateValue = dateTime.ToString("yyyy/MM/dd");
                if (hasWeekDay)
                {
                    dateValue += string.Format("({0})", ResourcesManager.WEEK_DAY_SHORT[(int)dateTime.DayOfWeek]);
                }
            }
            else if (yearMonth.Success)
            {
                dateValue = dateTime.Year + " " + ResourcesManager.MONTH_SHORT[dateTime.Month - 1];
            }
            else if (monthYear.Success)
            {
                dateValue = ResourcesManager.MONTH_SHORT[dateTime.Month - 1] + " " + dateTime.Year;
            }
            return dateValue;
        }
        public static DateTime ToDateTime(decimal date)
        {
            if (date == 0)
            {
                return DateTime.MinValue;
            }
            if (date == 99999999M)
            {
                return DateTime.MaxValue;
            }
            return DateTime.ParseExact(date.ToString(), "yyyyMMdd", null);
        }
        public static decimal GetCurrentDate()
        {
            return decimal.Parse(DateTime.Now.ToString("yyyyMMdd"));
        }

        public static bool IsInCurrentMonth(decimal date)
        {
            decimal currentDay = GetCurrentDate();
            decimal firstDayOfMonth = GetFirstDayOfMonth(currentDay);
            decimal lastDayOfMonth = GetLastDayOfMonth(currentDay);
            return firstDayOfMonth <= date && lastDayOfMonth >= date;
        }

        public static bool IsPassTime(decimal date)
        {
            return date < GetCurrentDate();
        }
        public static decimal? ToNullableDecimal(string value)
        {
            decimal data;
            if (string.IsNullOrEmpty(value) || !decimal.TryParse(value, out data))
            {
                return null;
            }
            return data;
        }
        public static decimal ToDecimal(string value)
        {
            return decimal.Parse(value);
        }
        public static decimal ToDecimal(DateTime dateTime)
        {
            return decimal.Parse(dateTime.ToString("yyyyMMdd"));
        }
        public static string GetFullName(string firstName, string lastName)
        {
            if (string.IsNullOrEmpty(firstName))
            {
                return lastName;
            }
            else if (string.IsNullOrEmpty(lastName))
            {
                return firstName;
            }
            string fullName = lastName + " " + firstName;
            return fullName.Replace("  ", " ");
        }
        public static decimal GetDateAsDecimal(string year, string month, string day)
        {
            return decimal.Parse(year + PaddingZero(month, 2) + PaddingZero(day, 2));
        }

        public static string PaddingZero(object value, int len)
        {
            if (value == null || value.ToString().Length == 0)
            {
                return null;
            }
            string str = value.ToString();
            while (str.Length < len)
            {
                str = "0" + str;
            }
            return str;
        }

        public static decimal AddTime(decimal from, decimal addTo)
        {
            decimal hrFrom = decimal.Floor(from / 100);
            decimal miFrom = (from - hrFrom * 100);
            decimal hrTo = decimal.Floor(addTo / 100);
            decimal miTo = (addTo - hrTo * 100);
            return (hrFrom + hrTo) * 60 + (miFrom + miTo);
        }
        public static decimal SubTime(decimal from, decimal to)
        {
            decimal hrFrom = decimal.Floor(from / 100);
            decimal miFrom = (from - hrFrom * 100);
            decimal hrTo = decimal.Floor(to / 100);
            decimal miTo = (to - hrTo * 100);
            return (hrFrom - hrTo) * 60 + (miFrom - miTo);
        }
        public static decimal DivideTime(decimal hhmm, int num)
        {
            decimal hr = decimal.Floor(hhmm / 100);
            decimal mi = (hhmm - hr * 100);
            decimal totalMi = hr * 60 + mi;
            decimal result = totalMi / num;
            return MinuteToHr(result);
        }
        public static string GetDirectory(string fileName)
        {
            // string path = Regex.Replace(fileName
            return null;
        }

        public static List<SelectItemModel> GetConstWorkDayTypeList()
        {
            List<SelectItemModel> list = new List<SelectItemModel>();
            string[] text = ResourcesManager.WORK_TYPE;
            int startNo = 0;
            foreach (var value in text)
            {
                string[] parts = value.Trim().Split('=');
                if (startNo == 0 && parts.Length == 2)
                {
                    startNo = int.Parse(parts[1].Trim());
                }
                else
                {
                    startNo++;
                }
                SelectItemModel dayType = new SelectItemModel()
                {
                    ItemCD = startNo.ToString(),
                    ItemValue = parts[0].Trim()
                };
                list.Add(dayType);
            }
            return list;
        }

        public static string FindWeekDay(string disp_date)
        {
            string value = Regex.Match(disp_date, "[(].*[)]$").Value.Replace("(", "").Replace(")", "");
            if (string.IsNullOrEmpty(value))
            {
                return string.Empty;
            }
            return value;
        }
        public static string FindDate(string disp_date)
        {
            string value = disp_date.Replace(Regex.Match(disp_date, "[(].*[)]$").Value, "");
            if (string.IsNullOrEmpty(value))
            {
                return string.Empty;
            }
            return value;
        }
        public static string ToDispHour(decimal? hour, bool displayZero = true)
        {
            if (hour == null)
            {
                return string.Empty;
            }
            if (hour == 0 && !displayZero)
            {
                return string.Empty;
            }
            if (hour > 2400)
            {
                hour = MinuteToHr(SubTime(hour.Value, 2400));
            }
            return hour.Value.ToString("00:00");
        }
        public static string ToDispHour(string hour, bool displayZero = true)
        {
            if (string.IsNullOrEmpty(hour))
            {
                return string.Empty;
            }
            return ToDispHour(ToNullableDecimal(hour), displayZero);
        }
        public static string ToDispMinute(decimal? minute, bool displayZero = true)
        {

            if (minute == null)
            {
                return string.Empty;
            }
            decimal hour = MinuteToHr(minute.Value);
            if (hour == 0 && !displayZero)
            {
                return string.Empty;
            }
            return hour.ToString("00:00");
        }
        public static string ToDispMinute(string minute, bool displayZero = true)
        {
            if (string.IsNullOrEmpty(minute))
            {
                return string.Empty;
            }
            return ToDispMinute(ToNullableDecimal(minute), displayZero);
        }

        public static string GetPosition(string remarks)
        {
            string pos = string.Empty;
            if (!string.IsNullOrEmpty(remarks) && remarks.Length > DBConstant.EMPL_TEAM_LEN)
            {
                pos = remarks.Substring(DBConstant.EMPL_TEAM_LEN).Trim();

            }
            else if (remarks != null)
            {
                pos = remarks.Trim();
            }
            return pos;
        }
        public static string GetTeam(string remarks)
        {
            string pos = string.Empty;
            if (!string.IsNullOrEmpty(remarks) && remarks.Length > DBConstant.EMPL_TEAM_LEN)
            {
                pos = remarks.Substring(0, DBConstant.EMPL_TEAM_LEN).Trim();

            }
            else if (remarks != null)
            {
                pos = remarks.Trim();
            }
            return pos;
        }


        public static object GetXlsValue(object value, bool isNotDisplayZero = true)
        {
            if (value == null || isNotDisplayZero && value.ToString().Equals("0"))
            {
                return string.Empty;
            }

            return value;
        }

        public static int GetExcelColor(string htmlColor)
        {
            System.Drawing.Color color = System.Drawing.ColorTranslator.FromHtml(htmlColor);
            return System.Drawing.ColorTranslator.ToOle(color);
        }
        public static string GetPresence(WorkDataModel workData)
        {
            string work = string.Empty;
            if (workData.IsHoliday && (string.IsNullOrEmpty(workData.Update_start_time) && string.IsNullOrEmpty(workData.Update_end_time)))
            {
                work = DBConstant.ABSENT_OFF;
            }
            else
            {
                if (workData.WorkingType == WorkingType.Present)
                {
                    switch ((int)workData.Work_type_no.Value)
                    {
                        case DBConstant.WORK_TYPE_NORMAL:
                        case DBConstant.WORK_TYPE_HOLIDAY_DUTY:
                            if (!string.IsNullOrEmpty(workData.Update_start_time) || !string.IsNullOrEmpty(workData.Update_end_time))
                            {
                                work = DBConstant.PRESENCE_PRESENT;
                            }

                            break;
                        case DBConstant.WORK_TYPE_HALF_PERMISSION:
                            if (!string.IsNullOrEmpty(workData.Update_start_time) || !string.IsNullOrEmpty(workData.Update_end_time))
                            {
                                work = DBConstant.ABSENT_HALF_PERMISSION;
                            }
                            break;
                        default:
                            work = string.Empty;
                            break;
                    }
                }
                else if (workData.WorkingType == WorkingType.Unknown)
                {
                    if (workData.IsOutOfExpiration)
                    {
                        work = DBConstant.CHAR_UNKNOWN;
                    }
                }
                else
                {
                    switch ((int)workData.Work_type_no.Value)
                    {
                        case DBConstant.WORK_TYPE_ANNUAL_LEAVE:
                            work = DBConstant.ABSENT_ANNUAL_LEAVE;
                            break;
                        case DBConstant.WORK_TYPE_AW_PERMISSION:
                            work = DBConstant.ABSENT_WITHOUT_PERMISSION;
                            break;
                        case DBConstant.WORK_TYPE_PERMISSION:
                            work = DBConstant.ABSENT_PERMISSION;
                            break;
                        case DBConstant.WORK_TYPE_SPECIAL_LEAVE:
                            work = DBConstant.ABSENT_SPECIAL_LEAVE;
                            break;
                        case DBConstant.WORK_TYPE_MATERNITY_LEAVE:
                            work = DBConstant.ABSENT_MATERNITY_LEAVE;
                            break;
                        default:
                            work = DBConstant.CHAR_UNKNOWN;
                            break;
                    }
                }
            }
            return work;
        }

        public static bool IsMiddleNightShift(TimeTableModel timeTable)
        {
            if (timeTable.Work_from >= 2200)
            {
                return true;
            }
            return false;
        }

        public static decimal GetOverTime(decimal workDate)
        {
            if (workDate > 1)
            {
                return workDate - 1;
            }
            return 0;
        }

        public static void SetBackColor(Range range, int startCell, int endCell, int color)
        {
            for (int i = startCell; i <= endCell; i++)
            {
                range.Columns[i].Interior.Color = color;
            }
        }

        public static void CreateEmpltyColumn(Worksheet sheet, int fromColum, int totalColumn)
        {

            for (int i = 0; i < totalColumn - 1; i++)
            {
                Range range = sheet.Range["B" + fromColum, "D" + fromColum].EntireColumn;
                range.Copy(Type.Missing);
                range.Insert(Microsoft.Office.Interop.Excel.XlInsertShiftDirection.xlShiftToRight, Type.Missing);
            }
        }

        public static void CreateEmptyRow(Worksheet sheet, int fromRow, int totalRow)
        {
            Range range = sheet.Range["A" + fromRow, Type.Missing].EntireRow;

            for (int i = 0; i < totalRow - 1; i++)
            {
                range.Copy(Type.Missing);
                range.Insert(Microsoft.Office.Interop.Excel.XlInsertShiftDirection.xlShiftDown, Type.Missing);
            }
            sheet.Select();
        }

        public static WorkingType GetWorkingType(WorkDataModel workData)
        {
            WorkingType workingType = WorkingType.Unknown;
            decimal sysDate = CommonUtil.GetCurrentDate();
            switch ((int)workData.Work_type_no)
            {
                case DBConstant.WORK_TYPE_HALF_PERMISSION:
                case DBConstant.WORK_TYPE_NORMAL:

                    if (workData.Work_date <= sysDate && !workData.IsHoliday)
                    {
                        workingType = WorkingType.Present;
                    }
                    else if (workData.IsHoliday)
                    {
                        workingType = WorkingType.Absent;
                    }
                    else
                    {
                        workingType = WorkingType.Unknown;
                    }
                    break;
                case DBConstant.WORK_TYPE_HOLIDAY_DUTY:
                    if (!string.IsNullOrEmpty(workData.Update_start_time) || !string.IsNullOrEmpty(workData.Update_end_time))
                    {
                        workingType = WorkingType.Present;
                    }
                    else
                    {
                        workingType = WorkingType.Absent;
                    }
                    break;
                default:
                    workingType = WorkingType.Absent;
                    break;
            }

            return workingType;
        }
        public static WorkingDayType GetWorkingDayType(WorkDataModel workData, EmployeeModel employee, List<HolidayModel> holidayList)
        {
            WorkingDayType workingType;
            HolidayModel holiday = null;
            if (employee.Use_flag_of_holiday == DBConstant.FLAG_OF_HOLIDAY_USE)
            {
                holiday = holidayList.Find(x => x.Holiday_date == workData.Work_date);
            }

            if (holiday == null)
            {
                if (workData.Work_day_type_no <= DBConstant.WORK_DAY_TYPE_NORMAL)
                {

                    workingType = WorkingDayType.Normal;

                }
                else if (workData.Work_day_type_no <= DBConstant.WORK_DAY_TYPE_NORMAL_HOLIDAY)
                {
                    workingType = WorkingDayType.NormalHoliday;
                }
                else
                {
                    workingType = WorkingDayType.NationalHoliday;
                }
            }
            else
            {
                if (holiday.National_holiday_flag == DBConstant.HOLIDAY_FLAG_NATIONAL)
                {
                    workingType = WorkingDayType.NationalHoliday;
                }
                else
                {
                    workingType = WorkingDayType.NormalHoliday;
                }
            }
            return workingType;
        }

        public static void DeleteDefaultSheet(Workbook workBook)
        {
            try
            {
                var sheet1 = workBook.Worksheets[Properties.Settings.Default.XLS_Default_Sheet_Name];
                if (sheet1 != null)
                {
                    sheet1.Delete();
                }
            }
            catch
            {
                //DO nothing
            }
            finally
            {
            }
        }

        public static void SetXlsPageSetup(Worksheet workSheet, XlPageOrientation orientation)
        {
            workSheet.PageSetup.Zoom = false;
            workSheet.PageSetup.TopMargin = 20;
            workSheet.PageSetup.BottomMargin = 30;
            workSheet.PageSetup.RightMargin = 30;
            workSheet.PageSetup.LeftMargin = 40;
            workSheet.PageSetup.FitToPagesWide = 1;
            workSheet.PageSetup.FitToPagesTall = 1000;
            workSheet.PageSetup.Orientation = orientation;
            workSheet.PageSetup.PaperSize = XlPaperSize.xlPaperA4;
            workSheet.PageSetup.CenterHorizontally = false;

        }
        public static void FillPersonalReportSheet(Worksheet employeeSheet, EmployeeModel employeeModel,
            List<WorkDataModel> workDataList, List<SelectItemModel> dayTypeList, List<TimeTableModel> timeTableList, decimal unit_minutes)
        {
            employeeSheet.Name = employeeModel.Employee_no;
            Microsoft.Office.Interop.Excel.Range employeeNameRange = employeeSheet.get_Range("G6", Type.Missing);
            Microsoft.Office.Interop.Excel.Range employeeCodeRange = employeeSheet.get_Range("G7", Type.Missing);
            Microsoft.Office.Interop.Excel.Range departmentRange = employeeSheet.get_Range("L6");
            Microsoft.Office.Interop.Excel.Range positionRange = employeeSheet.get_Range("L7");
            Microsoft.Office.Interop.Excel.Range teamRange = employeeSheet.get_Range("L8");
            Microsoft.Office.Interop.Excel.Range detailRange = employeeSheet.get_Range("A11", Type.Missing);

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
            decimal totalP = 0;
            int totalWP = 0;
            int totalSL = 0;
            int totalML = 0;
            employeeNameRange.Value = GetFullName(employeeModel.Emsize_first_name, employeeModel.Emsize_last_name); ;
            employeeCodeRange.Value = employeeModel.Employee_no;
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


            CreateEmptyRow(employeeSheet, startRow, workDataList.Count);
            int unitMinutes = 0;
            foreach (var workData in workDataList)
            {
                int columnIndex = 1;
                int aL = 0;
                decimal p = 0;
                int ml = 0;
                int wp = 0;
                int sl = 0;


                var work_type = dayTypeList.Find(x => decimal.Parse(x.ItemCD) == workData.Work_type_no);

                string work_type_name = "";


                if (!workData.IsOutOfExpiration)
                {
                    decimal? contact_time_tmp = CommonUtil.ToNullableDecimal(workData.Contract_time);
                    if (contact_time_tmp == null || contact_time_tmp == 0)
                    {
                        if (!workData.IsHoliday)
                        {
                            TimeTableModel timeTable = timeTableList.Find(x => x.Time_table_no == workData.Time_table_no);
                            if (timeTable != null)
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
                    }
                    else
                    {
                        contactTime += contact_time_tmp.Value;
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
                    if (workData.Work_day_type_no > DBConstant.WORK_DAY_TYPE_NORMAL)
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
                    case DBConstant.WORK_TYPE_HALF_PERMISSION:
                        p = DBConstant.WORKING_HALF_PERMISION;
                        totalP += DBConstant.WORKING_HALF_PERMISION;
                        break;
                    case DBConstant.WORK_TYPE_AW_PERMISSION:
                        wp = 1;
                        totalWP++;
                        break;
                    case DBConstant.WORK_TYPE_SPECIAL_LEAVE:
                        sl = 1;
                        totalSL++;
                        break;
                    case DBConstant.WORK_TYPE_MATERNITY_LEAVE:
                        ml = 1;
                        totalML++;
                        break;
                    default:
                        break;
                }

                Microsoft.Office.Interop.Excel.Range detailEditRange = employeeSheet.get_Range(string.Format(detailRowStart, startRow), string.Format(detailRowEnd, startRow));

                detailEditRange.Columns[columnIndex++].Value = CommonUtil.FindDate(workData.Work_date_dsp);
                detailEditRange.Columns[columnIndex++].Value = CommonUtil.FindWeekDay(workData.Work_date_dsp);
                if (workData.Work_type_no == DBConstant.WORK_TYPE_NORMAL && (!string.IsNullOrEmpty(workData.Update_start_time) || !string.IsNullOrEmpty(workData.Update_end_time))
                    || (workData.Work_type_no != DBConstant.WORK_TYPE_NORMAL))
                {
                    detailEditRange.Columns[columnIndex++].Value = work_type_name;
                    detailEditRange.Columns[columnIndex++].Value = workData.TimeTable;
                    detailEditRange.Columns[columnIndex++].Value = CommonUtil.ToDispHour(workData.Update_start_time);
                    detailEditRange.Columns[columnIndex++].Value = CommonUtil.ToDispHour(workData.Update_end_time);
                    detailEditRange.Columns[columnIndex++].Value = CommonUtil.ToDispMinute(workData.Working_time);
                    detailEditRange.Columns[columnIndex++].Value = CommonUtil.ToDispMinute(workData.Over_time);
                    detailEditRange.Columns[columnIndex++].Value = CommonUtil.ToDispMinute(workData.Late_night_time);
                    detailEditRange.Columns[columnIndex++].Value = CommonUtil.ToDispMinute(workData.Holiday_time);
                    detailEditRange.Columns[columnIndex++].Value = CommonUtil.ToDispMinute(workData.Holiday_late_night_time);
                }
                else
                {
                    columnIndex += 9;
                }
                detailEditRange.Columns[columnIndex++].Value = CommonUtil.GetXlsValue(aL);
                detailEditRange.Columns[columnIndex++].Value = CommonUtil.GetXlsValue(ml);
                detailEditRange.Columns[columnIndex++].Value = CommonUtil.GetXlsValue(p);
                detailEditRange.Columns[columnIndex++].Value = CommonUtil.GetXlsValue(wp);
                detailEditRange.Columns[columnIndex++].Value = CommonUtil.GetXlsValue(sl);
                detailEditRange.Columns[columnIndex++].Value = workData.Memo;

                startRow++;
            }
            // startRow += 1;
            Microsoft.Office.Interop.Excel.Range footer0Range = employeeSheet.get_Range("D" + startRow, "O" + startRow++);
            Microsoft.Office.Interop.Excel.Range footer1Range = employeeSheet.get_Range(string.Format(detailRowStart, startRow), string.Format(detailRowEnd, startRow++));
            Microsoft.Office.Interop.Excel.Range footer2Range = employeeSheet.get_Range(string.Format(detailRowStart, startRow), string.Format(detailRowEnd, startRow++));
            int col_contact_Ot = 4;
            int col_working_Mid = 7;
            int col_Holiday_H_Mid = 11;
            int col_P_Holiday_P_Mid = 15;
            int col_total_start_index = 4;
            footer0Range.Columns[col_total_start_index++].Value = CommonUtil.ToDispMinute(ConvertMinutesToOmissionHHMM(totalWorkingTime, unitMinutes), false);
            footer0Range.Columns[col_total_start_index++].Value = CommonUtil.ToDispMinute(ConvertMinutesToOmissionHHMM(totalOverTime, unitMinutes), false);
            footer0Range.Columns[col_total_start_index++].Value = CommonUtil.ToDispMinute(ConvertMinutesToOmissionHHMM(totalMidleNight, unitMinutes), false);
            footer0Range.Columns[col_total_start_index++].Value = CommonUtil.ToDispMinute(ConvertMinutesToOmissionHHMM(totalHolidayTime + totalPublicHolidayTime, unitMinutes), false);
            footer0Range.Columns[col_total_start_index++].Value = CommonUtil.ToDispMinute(ConvertMinutesToOmissionHHMM(totalHolidayMidleNigh, unitMinutes), false);
            footer0Range.Columns[col_total_start_index++].Value = CommonUtil.GetXlsValue(totalAL);
            footer0Range.Columns[col_total_start_index++].Value = CommonUtil.GetXlsValue(totalML);
            footer0Range.Columns[col_total_start_index++].Value = CommonUtil.GetXlsValue(totalP);
            footer0Range.Columns[col_total_start_index++].Value = CommonUtil.GetXlsValue(totalWP);
            footer0Range.Columns[col_total_start_index++].Value = CommonUtil.GetXlsValue(totalSL);

            footer1Range.Columns[col_contact_Ot].Value = CommonUtil.ToDispMinute(ConvertMinutesToOmissionHHMM(contactTime, unitMinutes), false);
            footer1Range.Columns[col_working_Mid].Value = CommonUtil.ToDispMinute(ConvertMinutesToOmissionHHMM(totalWorkingTime + totalHolidayTime + totalPublicHolidayTime, unitMinutes), false);
            footer1Range.Columns[col_Holiday_H_Mid].Value = CommonUtil.ToDispMinute(ConvertMinutesToOmissionHHMM(totalHolidayTime, unitMinutes), false);
            footer1Range.Columns[col_P_Holiday_P_Mid].Value = CommonUtil.ToDispMinute(ConvertMinutesToOmissionHHMM(totalPublicHolidayTime, unitMinutes), false);

            footer2Range.Columns[col_contact_Ot].Value = CommonUtil.ToDispMinute(ConvertMinutesToOmissionHHMM(totalOverTime, unitMinutes), false);
            footer2Range.Columns[col_working_Mid].Value = CommonUtil.ToDispMinute(ConvertMinutesToOmissionHHMM(totalMidleNight, unitMinutes), false);
            footer1Range.Columns[col_Holiday_H_Mid].Value = CommonUtil.ToDispMinute(ConvertMinutesToOmissionHHMM(totalHolidayMidleNigh, unitMinutes), false);
            footer2Range.Columns[col_P_Holiday_P_Mid].Value = CommonUtil.ToDispMinute(ConvertMinutesToOmissionHHMM(totalPublicMidleNightHolidayTime, unitMinutes), false);
        }

        public static List<T> GetSortedList<T>(List<T> list)
        {
            var colections = CollectionViewSource.GetDefaultView(list);
            List<T> tlist = new List<T>();
            foreach (T item in colections)
            {
                tlist.Add(item);
            }
            return tlist;
        }

        public static void ClearSortDirection<T>(List<T> list)
        {
            System.Windows.Data.ListCollectionView view = (System.Windows.Data.ListCollectionView)CollectionViewSource.GetDefaultView(list);
            if (view.IsAddingNew)
            {
                view.CommitNew();
            }
            if (view.IsEditingItem)
            {
                view.CommitEdit();
            }
        }
        public static void SaveXls(string fileName, Workbook workBook)
        {
            workBook.SaveAs(fileName);
            workBook.Saved = true;
        }

        public static string GetExcelColumnName(int column)
        {
            int input = column;
            string value = "";
            while (input > 0)
            {
                if (input <= 26)
                {
                    value = value + ((char)(input + 64)).ToString();
                    break;
                }
                else
                {
                    int kytucuoi = input % 26;
                    int kytudau = input / 26;
                    if (kytucuoi == 0)
                    {
                        kytudau--;
                    }

                    value = value + ((char)(kytudau + 64)).ToString();
                    input = (input - kytudau * 26);

                }
            }
            return value;
        }

        public static string GetTemplate(string templateName)
        {
            return Environment.CurrentDirectory + @"\Template\" + templateName;
        }

        public static string GetTimeShiftDsp(TimeTableModel timeTable)
        {
            return timeTable.Work_from.ToString("00:00") + "-" + timeTable.Work_to.ToString("00:00");
        }
        public static void ClearTemplate()
        {
            string templateDir = GetTemplate("");
            string[] files = Directory.GetFiles(templateDir);
            if (files != null && files.Length > 0)
            {
                string dailyTemplate = string.Format(Properties.Settings.Default.XLS_Daily_Report, "").Split('.')[0];
                string monthlyTemplate = string.Format(Properties.Settings.Default.XLS_Monthly_Report, "").Split('.')[0];
                string personTemplate = string.Format(Properties.Settings.Default.XLS_Personal_Report, "").Split('.')[0];
                string correctedTemplate = string.Format(Properties.Settings.Default.XLS_Correct_Night_Shift, "").Split('.')[0];
                foreach (string file in files)
                {
                    if (!(file.Contains(dailyTemplate)
                        || file.Contains(monthlyTemplate)
                        || file.Contains(correctedTemplate)
                        || file.Contains(personTemplate)))
                    {
                        File.Delete(file);
                    }
                }
            }
        }
        public static Workbook CreateWorkbook(Application excelApp, string fileName)
        {
            Application excelTemplateAp = new Application();
            excelTemplateAp.SheetsInNewWorkbook = 1;

            string fullFileName = GetTemplate(fileName);
            if (!File.Exists(fullFileName))
            {
                var workBook = excelTemplateAp.Workbooks.Add();
                if (workBook.ActiveSheet.Name != Properties.Settings.Default.XLS_Default_Sheet_Name)
                {
                    Properties.Settings.Default.XLS_Default_Sheet_Name = workBook.ActiveSheet.Name;
                    Properties.Settings.Default.Save();
                }
                workBook.SaveAs(fullFileName);
                workBook.Close();
                Marshal.ReleaseComObject(workBook);
                Marshal.ReleaseComObject(excelTemplateAp);
            }

            return excelApp.Workbooks.Add(fullFileName);

        }

        [DllImport("user32.dll")]
        private static extern uint SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int x, int y, int width, int height, int flags);
        public static uint SetWindowPosision(long hWnd, long hWndInsertAfter)
        {
            return SetWindowPos((IntPtr)hWnd, (IntPtr)hWndInsertAfter, 0, 0, 0, 0, 0);
        }

        public static uint SetWindowPosision(long hWnd, IntPtr hWndInsertAfter)
        {
            return SetWindowPos((IntPtr)hWnd, hWndInsertAfter, 0, 0, 0, 0, 0);
        }

        public static int ConvertHHMMToOmissionHHMM(int hhmm, int unitMinutes)
        {
            int num = Math.Abs(hhmm);
            bool flag = false;
            int length = num.ToString().Length;
            int hr = 0;
            int num2 = 1;
            if ((hhmm == 0) || (unitMinutes <= 0))
            {
                return hhmm;
            }
            if (hhmm < 0)
            {
                flag = true;
            }
            hr = (num / 100) * 100;
            int minute = num % 100;
            while ((unitMinutes * num2) <= 60)
            {
                if (minute < (unitMinutes * num2))
                {
                    num = hr + (unitMinutes * (num2 - 1));
                    break;
                }
                num2++;
            }
            if (flag)
            {
                return (num * -1);
            }
            return num;
        }


        public static decimal ConvertMinutesToOmissionHHMM(decimal minutes, int unitMinutes)
        {
            if ((minutes != 0) && (unitMinutes > 0))
            {
                return ConvertHHMMToOmissionHHMM((int)MinuteToHr(minutes), unitMinutes);
            }
            return minutes;
        }


        public static string Capitalize(string value)
        {
            if (string.IsNullOrEmpty(value))
            {
                return string.Empty;
            }
            string[] input = value.ToLower().Trim().Split(' ');
            string ouput = "";
            foreach (string srt in input)
            {
                ouput += srt[0].ToString().ToUpper();
                if (srt.Length > 1)
                {
                    ouput += srt.Substring(1);
                }
                ouput += " ";
            }
            return ouput.Trim();
        }

    }

    public enum WorkingDayType
    {
        Normal,
        NormalHoliday,
        NationalHoliday,
    }
    public enum WorkingType
    {
        Present,
        Absent,
        Unknown,
    }
}
