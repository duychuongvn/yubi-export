using System;
using System.Collections.Generic;
using System.ComponentModel.Composition;
using System.Data.Linq;
using System.Data.Linq.SqlClient;
using System.Linq;
using YUBI_TOOL.Common;
using YUBI_TOOL.Dao;
using YUBI_TOOL.Model;

namespace YUBI_TOOL.Service
{
    [Export(typeof(IWorkDataService))]
    public class WorkDataService : IWorkDataService
    {
        public void UpdateWorkDataList(List<WorkDataModel> workDataList)
        {
            decimal sysDate = Common.CommonUtil.GetCurrentDate();
            YubitaroDataContext context = new YubitaroDataContext();
            foreach (var workDataModel in workDataList)
            {
                var workData = context.WORK_DATAs.Single(x => x.EMPLOYEE_NO == workDataModel.Employee_no
                    && x.COMPANY_NO == workDataModel.Company_no
                    && x.WORK_DATE == workDataModel.Work_date);
                if (workData != null)
                {
                    workData.ABSENCE_DAYS = workDataModel.Absence_days;
                    workData.BEING_LATE_DAYS = workDataModel.Being_late_days;
                    workData.BEING_LATE_TIME = CommonUtil.ToNullableDecimal(workDataModel.Being_late_time);
                    workData.COMPENSATORY_DAY_OFF = workDataModel.Compensatory_day_off;
                    workData.CONTRACT_TIME = CommonUtil.ToNullableDecimal(workDataModel.Contract_time);
                    workData.DILIGENCE_INDOLENCE_POINT = workDataModel.Diligence_indolence_point;
                    workData.END_TIME = CommonUtil.ToNullableDecimal(workDataModel.End_time);
                    workData.HOLIDAY_DAYS = workDataModel.Holiday_days;
                    workData.HOLIDAY_LATE_NIGHT_TIME = CommonUtil.ToNullableDecimal(workDataModel.Holiday_late_night_time);
                    workData.HOLIDAY_TIME = CommonUtil.ToNullableDecimal(workDataModel.Holiday_time);
                    workData.LATE_NIGHT_TIME = CommonUtil.ToNullableDecimal(workDataModel.Late_night_time);
                    workData.LEAVING_EARLY_DAYS = workDataModel.Leaving_early_days;
                    workData.LEAVING_EARLY_TIME = CommonUtil.ToNullableDecimal(workDataModel.Leaving_early_time);
                    workData.MEMO = workDataModel.Memo;
                    workData.OVER_TIME = CommonUtil.ToNullableDecimal(workDataModel.Over_time);
                    workData.PAID_VACATION_DAYS = workDataModel.Paid_vacation_days;
                    workData.PAID_VACATION_TIME = CommonUtil.ToNullableDecimal(workDataModel.Paid_vacation_time);
                    workData.REST_TIME = CommonUtil.ToNullableDecimal(workDataModel.Rest_time);
                    workData.SPECIAL_HOLIDAYS = workDataModel.Special_holidays;
                    workData.START_TIME = CommonUtil.ToNullableDecimal(workDataModel.Start_time);
                    workData.TIME_TABLE_NO = workDataModel.Time_table_no;
                    workData.UPDATE_DATE_TIME = DateTime.Now;
                    workData.UPDATE_END_TIME = CommonUtil.ToNullableDecimal(workDataModel.Update_end_time);
                    workData.UPDATE_START_TIME = CommonUtil.ToNullableDecimal(workDataModel.Update_start_time);
                    workData.WORK_DAY_TYPE_NO = workDataModel.Work_day_type_no;
                    workData.WORK_DAYS = workDataModel.Work_days;
                    workData.WORK_TYPE_NO = workDataModel.Work_type_no;
                    workData.WORKING_TIME = CommonUtil.ToNullableDecimal(workDataModel.Working_time);
                }

            }
            try
            {
                context.SubmitChanges();
            }
            catch (ChangeConflictException)
            {

            }
        }
        public WorkDataModel SearchWorkData(decimal companyNo, string employeeNo, decimal workDate)
        {
            var context = Dao.DaoHelper.GetContext();
            var result = (from wd in context.WORK_DATAs
                          where wd.COMPANY_NO == companyNo
                          && wd.EMPLOYEE_NO == employeeNo
                          && wd.WORK_DATE == workDate
                          select wd).FirstOrDefault();
            if (result == null)
            {
                return null;
            }
            WorkDataModel workDataModel = new WorkDataModel
            {
                Absence_days = result.ABSENCE_DAYS,
                Being_late_days = result.BEING_LATE_DAYS,
                Being_late_time = CommonUtil.ToString(result.BEING_LATE_TIME),
                Company_no = result.COMPANY_NO,
                Compensatory_day_off = result.COMPENSATORY_DAY_OFF,
                Contract_time = CommonUtil.ToString(result.CONTRACT_TIME),
                Create_date_time = result.CREATE_DATE_TIME,
                Diligence_indolence_point = result.DILIGENCE_INDOLENCE_POINT,
                Employee_no = result.EMPLOYEE_NO,
                End_time = CommonUtil.ToString(result.END_TIME),
                Holiday_days = result.HOLIDAY_DAYS,
                Holiday_late_night_time = CommonUtil.ToString(result.HOLIDAY_LATE_NIGHT_TIME),
                Holiday_time = CommonUtil.ToString(result.HOLIDAY_TIME),
                Late_night_time = CommonUtil.ToString(result.LATE_NIGHT_TIME),
                Leaving_early_days = result.LEAVING_EARLY_DAYS,
                Leaving_early_time = CommonUtil.ToString(result.LEAVING_EARLY_TIME),
                Memo = result.MEMO,
                Over_time = CommonUtil.ToString(result.OVER_TIME),
                Paid_vacation_days = result.PAID_VACATION_DAYS,
                Paid_vacation_time = CommonUtil.ToString(result.PAID_VACATION_TIME),
                Rest_time = CommonUtil.ToString(result.REST_TIME),
                Special_holidays = result.SPECIAL_HOLIDAYS,
                Start_time = CommonUtil.ToString(result.START_TIME),
                Time_table_no = result.TIME_TABLE_NO,
                Update_date_time = result.UPDATE_DATE_TIME,
                Update_end_time = CommonUtil.ToString(result.UPDATE_END_TIME),

                Update_start_time = CommonUtil.ToString(result.UPDATE_START_TIME),
                Work_date = result.WORK_DATE,
                Work_day_type_no = result.WORK_DAY_TYPE_NO,
                Work_days = result.WORK_DAYS,
                Work_type_no = result.WORK_TYPE_NO,
                Working_time = CommonUtil.ToString(result.WORKING_TIME),
            };
            return workDataModel;
        }
        public List<WorkDataModel> SearchWorkDataListNightShift(decimal companyNo, decimal dateForm, decimal dateTo)
        {
            decimal endDayOfMonth = CommonUtil.GetLastDayOfMonth(dateTo);
            DateTime endDayOfMonthDateTime = CommonUtil.ToDateTime(endDayOfMonth);
            decimal firstDayOfNextMonth = CommonUtil.ToDecimal(endDayOfMonthDateTime.AddDays(1));
            List<WorkDataModel> workDataList = new List<WorkDataModel>();
            var context = Dao.DaoHelper.GetContext();
            var results = from wd1 in context.WORK_DATAs
                          join tb in context.TIME_TABLEs
                          on wd1.TIME_TABLE_NO equals tb.TIME_TABLE_NO
                          where tb.WORK_FROM <= 2400M && (tb.WORK_TO > 2400 || tb.WORK_TO < tb.WORK_FROM)
                          && wd1.WORK_DATE >= dateForm
                          && wd1.WORK_DATE <= dateTo
                          && wd1.START_TIME != null
                          && wd1.UPDATE_START_TIME != null
                          && wd1.END_TIME == null
                          && wd1.UPDATE_END_TIME == null
                          && wd1.COMPANY_NO == companyNo
                          && context.WORK_DATAs.Any(
                            x => x.COMPANY_NO == wd1.COMPANY_NO
                                && x.EMPLOYEE_NO == wd1.EMPLOYEE_NO
                                && (x.WORK_DATE == wd1.WORK_DATE + 1
                                    || (wd1.WORK_DATE == endDayOfMonth && x.WORK_DATE == firstDayOfNextMonth))
                                && x.START_TIME == null
                                && x.UPDATE_START_TIME == null
                                && x.END_TIME != null
                                && x.UPDATE_END_TIME != null
                                && (x.WORKING_TIME == null || x.WORKING_TIME == 0))
                          select wd1;
            foreach (var result in results)
            {
                WorkDataModel workData = new WorkDataModel
                {

                    Absence_days = result.ABSENCE_DAYS,
                    Being_late_days = result.BEING_LATE_DAYS,
                    Being_late_time = CommonUtil.ToString(result.BEING_LATE_TIME),
                    Company_no = result.COMPANY_NO,
                    Compensatory_day_off = result.COMPENSATORY_DAY_OFF,
                    Contract_time = CommonUtil.ToString(result.CONTRACT_TIME),
                    Create_date_time = result.CREATE_DATE_TIME,
                    Diligence_indolence_point = result.DILIGENCE_INDOLENCE_POINT,
                    Employee_no = result.EMPLOYEE_NO,
                    End_time = CommonUtil.ToString(result.END_TIME),
                    Holiday_days = result.HOLIDAY_DAYS,
                    Holiday_late_night_time = CommonUtil.ToString(result.HOLIDAY_LATE_NIGHT_TIME),
                    Holiday_time = CommonUtil.ToString(result.HOLIDAY_TIME),
                    Late_night_time = CommonUtil.ToString(result.LATE_NIGHT_TIME),
                    Leaving_early_days = result.LEAVING_EARLY_DAYS,
                    Leaving_early_time = CommonUtil.ToString(result.LEAVING_EARLY_TIME),
                    Memo = result.MEMO,
                    Over_time = CommonUtil.ToString(result.OVER_TIME),
                    Paid_vacation_days = result.PAID_VACATION_DAYS,
                    Paid_vacation_time = CommonUtil.ToString(result.PAID_VACATION_TIME),
                    Rest_time = CommonUtil.ToString(result.REST_TIME),
                    Special_holidays = result.SPECIAL_HOLIDAYS,
                    Start_time = CommonUtil.ToString(result.START_TIME),
                    Time_table_no = result.TIME_TABLE_NO,
                    Update_date_time = result.UPDATE_DATE_TIME,
                    Update_end_time = CommonUtil.ToString(result.UPDATE_END_TIME),

                    Update_start_time = CommonUtil.ToString(result.UPDATE_START_TIME),
                    Work_date = result.WORK_DATE,
                    Work_day_type_no = result.WORK_DAY_TYPE_NO,
                    Work_days = result.WORK_DAYS,
                    Work_type_no = result.WORK_TYPE_NO,
                    Working_time = CommonUtil.ToString(result.WORKING_TIME),
                };
                workDataList.Add(workData);
            }
            return workDataList;

        }
        public List<WorkDataModel> SearchWorkDataList(decimal companyNo, decimal postNo, string employeeName, decimal dateForm, decimal dateTo)
        {
            List<WorkDataModel> workDataList = new List<WorkDataModel>();
            decimal firstDayOfMonth = CommonUtil.GetFirstDayOfMonth(dateForm);
            decimal lastDayOfMonth = CommonUtil.GetLastDayOfMonth(dateForm);
            var context = Dao.DaoHelper.GetContext();
            var results = from data in context.WORK_DATAs
                          join employee in context.EMPLOYEEs
                          on data.EMPLOYEE_NO equals employee.EMPLOYEE_NO into e1
                          from employee in
                              (from e2 in e1
                               where
                                   e2.EXPIRATION_FROM <= data.WORK_DATE && e2.EXPIRATION_TO >= data.WORK_DATE
                                   && e2.COMPANY_NO == data.COMPANY_NO
                               select e2).DefaultIfEmpty()
                          //from employee in e1.DefaultIfEmpty().Where(x => x.EXPIRATION_FROM <= data.WORK_DATE && x.EXPIRATION_TO >= data.WORK_DATE)
                          join post in context.POSTs
                          on employee.POST_NO equals post.POST_NO into p1
                          from post in
                              (from p2 in p1 where p2.COMPANY_NO == employee.COMPANY_NO select p2).DefaultIfEmpty()
                          //from post in p1.DefaultIfEmpty()
                          where data.WORK_DATE >= dateForm && data.WORK_DATE <= dateTo
                         && data.COMPANY_NO == companyNo
                         //&& employee.COMPANY_NO == companyNo
                         //&& post.COMPANY_NO == companyNo
                         && (SqlMethods.Like(employee.EMPLOYEE_NO, string.Format("%{0}%", employeeName))
                           || SqlMethods.Like(employee.EMSIZE_FIRST_NAME + employee.EMSIZE_LAST_NAME, string.Format("%{0}%", employeeName))
                           )
                          orderby employee.POST_NO ascending
                          orderby employee.EMPLOYEE_NO ascending
                          select new
                          {
                              data,
                              post.POST_NAME,
                              post.POST_NO,
                              employee.EMSIZE_FIRST_NAME,
                              employee.EMSIZE_LAST_NAME,
                              employee.REMARKS
                          };

            if (postNo > 0)
            {
                results = results.Where(x => x.POST_NO == postNo);
            }
            foreach (var result in results)
            {
                WorkDataModel workData = new WorkDataModel
                {
                    EmployeeName = Common.CommonUtil.GetFullName(result.EMSIZE_FIRST_NAME, result.EMSIZE_LAST_NAME),
                    Post_name = result.POST_NAME,
                    Post_no = (result.POST_NO == null) ? 0 : result.POST_NO.Value,
                    Absence_days = result.data.ABSENCE_DAYS,
                    Being_late_days = result.data.BEING_LATE_DAYS,
                    Being_late_time = CommonUtil.ToString(result.data.BEING_LATE_TIME),
                    Company_no = result.data.COMPANY_NO,
                    Compensatory_day_off = result.data.COMPENSATORY_DAY_OFF,
                    Contract_time = CommonUtil.ToString(result.data.CONTRACT_TIME),
                    Create_date_time = result.data.CREATE_DATE_TIME,
                    Diligence_indolence_point = result.data.DILIGENCE_INDOLENCE_POINT,
                    Employee_no = result.data.EMPLOYEE_NO,
                    End_time = CommonUtil.ToString(result.data.END_TIME),
                    Holiday_days = result.data.HOLIDAY_DAYS,
                    Holiday_late_night_time = CommonUtil.ToString(result.data.HOLIDAY_LATE_NIGHT_TIME),
                    Holiday_time = CommonUtil.ToString(result.data.HOLIDAY_TIME),
                    Late_night_time = CommonUtil.ToString(result.data.LATE_NIGHT_TIME),
                    Leaving_early_days = result.data.LEAVING_EARLY_DAYS,
                    Leaving_early_time = CommonUtil.ToString(result.data.LEAVING_EARLY_TIME),
                    Memo = result.data.MEMO,
                    Over_time = CommonUtil.ToString(result.data.OVER_TIME),
                    Paid_vacation_days = result.data.PAID_VACATION_DAYS,
                    Paid_vacation_time = CommonUtil.ToString(result.data.PAID_VACATION_TIME),
                    Rest_time = CommonUtil.ToString(result.data.REST_TIME),
                    Special_holidays = result.data.SPECIAL_HOLIDAYS,
                    Start_time = CommonUtil.ToString(result.data.START_TIME),
                    Time_table_no = result.data.TIME_TABLE_NO,
                    Update_date_time = result.data.UPDATE_DATE_TIME,
                    Update_end_time = CommonUtil.ToString(result.data.UPDATE_END_TIME),

                    Update_start_time = CommonUtil.ToString(result.data.UPDATE_START_TIME),
                    Work_date = result.data.WORK_DATE,
                    Work_day_type_no = result.data.WORK_DAY_TYPE_NO,
                    Work_days = result.data.WORK_DAYS,
                    Work_type_no = result.data.WORK_TYPE_NO,
                    Working_time = CommonUtil.ToString(result.data.WORKING_TIME),
                    Employee_remarks = result.REMARKS
                };
                workDataList.Add(workData);
            }
            return workDataList;
        }

        public List<WorkDataModel> SearchWorkDataListByEmployee(decimal companyNo, decimal postNo, string employeeNo, decimal dateForm, decimal dateTo)
        {
            List<WorkDataModel> workDataList = new List<WorkDataModel>();
            var context = Dao.DaoHelper.GetContext();
            var results = from data in context.WORK_DATAs
                          join employee in context.EMPLOYEEs
                          on data.EMPLOYEE_NO equals employee.EMPLOYEE_NO into e1
                          //from employee in e1.DefaultIfEmpty()
                          from employee in
                              (from e2 in e1
                               where e2.EXPIRATION_FROM <= data.WORK_DATE && e2.EXPIRATION_TO >= data.WORK_DATE
                               && e2.COMPANY_NO == data.COMPANY_NO
                               select e2).DefaultIfEmpty()

                          join post in context.POSTs
                          on employee.POST_NO equals post.POST_NO into p1
                          from post in
                              (from p2 in p1 where p2.COMPANY_NO == employee.COMPANY_NO select p2).DefaultIfEmpty()
                          // from post in p1.DefaultIfEmpty()
                          where data.WORK_DATE >= dateForm && data.WORK_DATE <= dateTo
                         && data.COMPANY_NO == companyNo
                              //&& employee.COMPANY_NO == companyNo
                              //&& post.COMPANY_NO == companyNo
                         && data.EMPLOYEE_NO == employeeNo
                          //&& (employee.EXPIRATION_FROM <= data.WORK_DATE && employee.EXPIRATION_TO >= data.WORK_DATE)
                          orderby data.WORK_DATE ascending
                          select new
                          {
                              data,
                              post.POST_NAME,
                              post.POST_NO,
                              employee.EMSIZE_FIRST_NAME,
                              employee.EMSIZE_LAST_NAME,
                              employee.REMARKS

                          };
            foreach (var result in results)
            {
                WorkDataModel workData = new WorkDataModel
                {
                    EmployeeName = Common.CommonUtil.GetFullName(result.EMSIZE_FIRST_NAME, result.EMSIZE_LAST_NAME),
                    Post_name = result.POST_NAME,
                    Post_no = result.POST_NO == null ? 0 : result.POST_NO.Value,
                    Absence_days = result.data.ABSENCE_DAYS,
                    Being_late_days = result.data.BEING_LATE_DAYS,
                    Being_late_time = CommonUtil.ToString(result.data.BEING_LATE_TIME),
                    Company_no = result.data.COMPANY_NO,
                    Compensatory_day_off = result.data.COMPENSATORY_DAY_OFF,
                    Contract_time = CommonUtil.ToString(result.data.CONTRACT_TIME),
                    Create_date_time = result.data.CREATE_DATE_TIME,
                    Diligence_indolence_point = result.data.DILIGENCE_INDOLENCE_POINT,
                    Employee_no = result.data.EMPLOYEE_NO,
                    End_time = CommonUtil.ToString(result.data.END_TIME),
                    Holiday_days = result.data.HOLIDAY_DAYS,
                    Holiday_late_night_time = CommonUtil.ToString(result.data.HOLIDAY_LATE_NIGHT_TIME),
                    Holiday_time = CommonUtil.ToString(result.data.HOLIDAY_TIME),
                    Late_night_time = CommonUtil.ToString(result.data.LATE_NIGHT_TIME),
                    Leaving_early_days = result.data.LEAVING_EARLY_DAYS,
                    Leaving_early_time = CommonUtil.ToString(result.data.LEAVING_EARLY_TIME),
                    Memo = result.data.MEMO,
                    Over_time = CommonUtil.ToString(result.data.OVER_TIME),
                    Paid_vacation_days = result.data.PAID_VACATION_DAYS,
                    Paid_vacation_time = CommonUtil.ToString(result.data.PAID_VACATION_TIME),
                    Rest_time = CommonUtil.ToString(result.data.REST_TIME),
                    Special_holidays = result.data.SPECIAL_HOLIDAYS,
                    Start_time = CommonUtil.ToString(result.data.START_TIME),
                    Time_table_no = result.data.TIME_TABLE_NO,
                    Update_date_time = result.data.UPDATE_DATE_TIME,
                    Update_end_time = CommonUtil.ToString(result.data.UPDATE_END_TIME),

                    Update_start_time = CommonUtil.ToString(result.data.UPDATE_START_TIME),
                    Work_date = result.data.WORK_DATE,
                    Work_day_type_no = result.data.WORK_DAY_TYPE_NO,
                    Work_days = result.data.WORK_DAYS,
                    Work_type_no = result.data.WORK_TYPE_NO,
                    Working_time = CommonUtil.ToString(result.data.WORKING_TIME),
                    Employee_remarks = result.REMARKS
                };
                workDataList.Add(workData);
            }

            return workDataList;
        }

        public void CreateWorkDataInMonth(EmployeeModel employee, decimal yearMonthDate)
        {
            decimal firstDayOfMonth = CommonUtil.GetFirstDayOfMonth(yearMonthDate);
            decimal lastDayOfMonth = CommonUtil.GetLastDayOfMonth(yearMonthDate);
            // var workByEmployeeList = this.SearchWorkDataListByEmployee(employee.Company_no, employee.Post_no, employee.Employee_no, firstDayOfMonth, lastDayOfMonth);
            var context = DaoHelper.GetContext();
            bool hasAddNewData = false;
            if (employee.Expiration_from <= firstDayOfMonth && employee.Expiration_to >= firstDayOfMonth
                || employee.Expiration_from <= lastDayOfMonth && employee.Expiration_to >= lastDayOfMonth)
            {

                for (decimal i = firstDayOfMonth; i <= lastDayOfMonth; i++)
                {
                    var countByWorkDate = context.WORK_DATAs.Count(x => x.COMPANY_NO == employee.Company_no
                        && x.EMPLOYEE_NO == employee.Employee_no && x.WORK_DATE == i);
                    WorkDataModel work = SearchWorkData(employee.Company_no, employee.Employee_no, i);
                    if (countByWorkDate == 0)
                    {

                        WORK_DATA workData = new WORK_DATA()
                        {
                            COMPANY_NO = employee.Company_no,
                            EMPLOYEE_NO = employee.Employee_no,
                            WORK_DATE = i,
                            WORK_DAY_TYPE_NO = DBConstant.WORK_DAY_TYPE_NORMAL,
                            WORK_TYPE_NO = DBConstant.WORK_TYPE_NORMAL,
                            TIME_TABLE_NO = employee.Time_table_no,
                            CREATE_DATE_TIME = DateTime.Now,
                            UPDATE_DATE_TIME = DateTime.Now,
                        };
                        context.WORK_DATAs.InsertOnSubmit(workData);
                        hasAddNewData = true;
                    }
                }
            }
            if (hasAddNewData)
            {
                context.SubmitChanges();
            }
        }
    }
}
