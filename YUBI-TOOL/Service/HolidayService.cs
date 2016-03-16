using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.ComponentModel.Composition;
using YUBI_TOOL.Model;

namespace YUBI_TOOL.Service
{
    [Export(typeof(IHolidayService))]
    public class HolidayService : IHolidayService
    {
        public List<HolidayModel> SearchHolidayList(decimal companyNo, decimal holidayFrom, decimal holidayTo)
        {
            List<HolidayModel> holidayList = new List<HolidayModel>();

            decimal sysDate = Common.CommonUtil.GetCurrentDate();
            var context = Dao.DaoHelper.GetContext();
            var results = from holiday in context.HOLIDAYs
                          where holiday.COMPANY_NO == companyNo
                          && holiday.HOLIDAY_DATE >= holidayFrom && holiday.HOLIDAY_DATE <= holidayTo
                          select holiday;
            foreach (var result in results)
            {
                HolidayModel holiday = new HolidayModel()
                {
                    Company_no = result.COMPANY_NO,
                    Create_date_time = result.CREATE_DATE_TIME,
                    Holiday_date = result.HOLIDAY_DATE,
                    National_holiday_flag = result.NATIONAL_HOLIDAY_FLAG,
                    Remarks = result.REMARKS,
                    Update_date_time = result.UPDATE_DATE_TIME
                };
                holidayList.Add(holiday);
            }

            return holidayList;
        }
    }
}
