using System.Collections.Generic;
using YUBI_TOOL.Model;

namespace YUBI_TOOL.Service
{
    public interface IHolidayService
    {
        List<HolidayModel> SearchHolidayList(decimal companyNo, decimal holidayFrom, decimal holidayTo);
    }
}
