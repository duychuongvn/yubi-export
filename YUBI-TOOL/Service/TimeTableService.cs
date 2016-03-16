using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.ComponentModel.Composition;
using YUBI_TOOL.Model;

namespace YUBI_TOOL.Service
{
    [Export(typeof(ITimeTableService))]
    public class TimeTableService : ITimeTableService
    {
        public List<TimeTableModel> SearchTimeTableList()
        {
            List<TimeTableModel> timeTableList = new List<TimeTableModel>();
            decimal sysDate = Common.CommonUtil.GetCurrentDate();
            var context = Dao.DaoHelper.GetContext();
            var results = from timeTable in context.TIME_TABLEs
                          where timeTable.EXPIRATION_FROM <= sysDate
                          && timeTable.EXPIRATION_TO >= sysDate
                          && timeTable.STATUS == Common.DBConstant.STATUS_ADD
                          orderby timeTable.TIME_TABLE_NO ascending
                          select timeTable;
            foreach (var result in results)
            {
                TimeTableModel timeTable = new TimeTableModel()
                {
                    Abbreviation = result.ABBREVIATION,
                    Coretime_from = result.CORETIME_FROM,
                    Coretime_to = result.CORETIME_TO,
                    Create_date_time = result.CREATE_DATE_TIME,
                    Delimitation = result.DELIMITATION,
                    Expiration_from = result.EXPIRATION_FROM,
                    Expiration_to = result.EXPIRATION_TO,
                    Midnight_work_from = result.MIDNIGHT_WORK_FROM,
                    Midnight_work_to = result.MIDNIGHT_WORK_TO,
                    Over_unit_minutes = result.OVER_UNIT_MINUTES,
                    Rest1_from = result.REST1_FROM,
                    Rest1_to = result.REST1_TO,
                    Rest2_from = result.REST2_FROM,
                    Rest2_to = result.REST2_TO,
                    Rest3_from = result.REST3_FROM,
                    Rest3_to = result.REST3_TO,
                    Rest4_from = result.REST4_FROM,
                    Rest4_to = result.REST4_TO,
                    Rest5_from = result.REST5_FROM,
                    Rest5_to = result.REST5_TO,
                    Rest6_from = result.REST6_FROM,
                    Rest6_to = result.REST6_TO,
                    Rest7_from = result.REST7_FROM,
                    Rest7_to = result.REST7_TO,
                    Rest8_from = result.REST8_FROM,
                    Rest8_to = result.REST8_TO,
                    Rest9_from = result.REST9_FROM,
                    Rest9_to = result.REST9_TO,
                    Status = result.STATUS,
                    Time_table_name = result.TIME_TABLE_NAME,
                    Time_table_no = result.TIME_TABLE_NO,
                    Unit_minutes = result.UNIT_MINUTES,
                    Update_date_time = result.UPDATE_DATE_TIME,
                    Work_from = result.WORK_FROM,
                    Work_to = result.WORK_TO,
                };
                timeTableList.Add(timeTable);
            }
            return timeTableList;
        }
    }
}
