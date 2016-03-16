
namespace YUBI_TOOL.Model
{
    public class HolidayModel : ModelBase
    {
        private decimal company_no;
        private decimal? holiday_date;
        private decimal national_holiday_flag;
        private string remarks;

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

        public decimal? Holiday_date
        {
            get
            {
                return holiday_date;
            }
            set
            {
                if (holiday_date != value)
                {
                    holiday_date = value;
                    NotifyOfPropertyChange(() => Holiday_date);
                }
            }
        }

        public decimal National_holiday_flag
        {
            get
            {
                return national_holiday_flag;
            }
            set
            {
                if (national_holiday_flag != value)
                {
                    national_holiday_flag = value;
                    NotifyOfPropertyChange(() => National_holiday_flag);
                }
            }
        }

        public string Remarks
        {
            get
            {
                return remarks;
            }
            set
            {
                if (remarks != value)
                {
                    remarks = value;
                    NotifyOfPropertyChange(() => Remarks);
                }
            }
        }
    }
}
