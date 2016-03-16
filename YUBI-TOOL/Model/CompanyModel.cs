
namespace YUBI_TOOL.Model
{
    public class CompanyModel : ModelBase
    {
        private string company_no;
        private decimal? expiration_from;
        private decimal? expiration_to;
        private string company_name;
        private decimal? fiscal_year_from;
        private decimal? paid_vacation_time_days;
        private int working_minutes;
        private int status;


        public string Company_no
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

        public decimal? Expiration_from
        {
            get
            {
                return expiration_from;
            }
            set
            {
                if (expiration_from != value)
                {
                    expiration_from = value;
                    NotifyOfPropertyChange(() => Expiration_from);
                }
            }
        }

        public decimal? Expiration_to
        {
            get
            {
                return expiration_to;
            }
            set
            {
                if (expiration_to != value)
                {
                    expiration_to = value;
                    NotifyOfPropertyChange(() => Expiration_to);
                }
            }
        }

        public string Company_name
        {
            get
            {
                return company_name;
            }
            set
            {
                if (company_name != value)
                {
                    company_name = value;
                    NotifyOfPropertyChange(() => Company_name);
                }
            }
        }

        public decimal? Fiscal_year_from
        {
            get
            {
                return fiscal_year_from;
            }
            set
            {
                if (fiscal_year_from != value)
                {
                    fiscal_year_from = value;
                    NotifyOfPropertyChange(() => Fiscal_year_from);
                }
            }
        }

        public decimal? Paid_vacation_time_days
        {
            get
            {
                return paid_vacation_time_days;
            }
            set
            {
                if (paid_vacation_time_days != value)
                {
                    paid_vacation_time_days = value;
                    NotifyOfPropertyChange(() => Paid_vacation_time_days);
                }
            }
        }

        public int Working_minutes
        {
            get
            {
                return working_minutes;
            }
            set
            {
                if (working_minutes != value)
                {
                    working_minutes = value;
                    NotifyOfPropertyChange(() => Working_minutes);
                }
            }
        }

        public int Status
        {
            get
            {
                return status;
            }
            set
            {
                if (status != value)
                {
                    status = value;
                    NotifyOfPropertyChange(() => Status);
                }
            }
        }

     
    }
}
