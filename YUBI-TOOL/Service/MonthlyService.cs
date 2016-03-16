using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.ComponentModel.Composition;
using YUBI_TOOL.Model;

namespace YUBI_TOOL.Service
{
    [Export(typeof(IMonthlyService))]
    public class MonthlyService : IMonthlyService
    {
        public MonthlyModel GetMonthly(decimal company_no)
        {
            MonthlyModel monthlyModel = null;
            var context = Dao.DaoHelper.GetContext();
            var result = (from m in context.MONTHLies
                          where m.COMPANY_NO == company_no
                          select m).FirstOrDefault();
            if (result != null)
            {
                monthlyModel = new MonthlyModel()
                {
                    Company_no = company_no,
                    Create_date_time = result.CREATE_DATE_TIME,
                    Cutoff_day = result.CUTOFF_DAY,
                    Expiration_from = result.EXPIRATION_FROM,
                    Expiration_to = result.EXPIRATION_TO,
                    Unit_minutes = result.UNIT_MINUTES,
                    Update_date_time = result.UPDATE_DATE_TIME,
                };

            }

            return monthlyModel;
        }
    }
}
