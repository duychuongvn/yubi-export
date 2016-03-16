
namespace YUBI_TOOL.Model
{
    public class EmployeeModel : ModelBase
    {
        private decimal id;
        private decimal expiration_from;
        private decimal expiration_to;
        private string employee_no;
        private string login_password;
        private decimal company_no;
        private decimal post_no;
        private string emsize_last_name;
        private string emsize_first_name;
        private string alphabet_last_name;
        private string alphabet_first_name;
        private decimal? time_table_no;
        private decimal use_flag_of_holiday;
        private decimal? etirement_date;
        private string remarks;
        private decimal status;
        private string post_name;

        #region get/set

        public decimal Id
        {
            get
            {
                return id;
            }
            set
            {
                if (id != value)
                {
                    id = value;
                    NotifyOfPropertyChange(() => Id);
                }
            }
        }

        public decimal Expiration_from
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

        public decimal Expiration_to
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

        public string Employee_no
        {
            get
            {
                return employee_no;
            }
            set
            {
                if (employee_no != value)
                {
                    employee_no = value;
                    NotifyOfPropertyChange(() => Employee_no);
                }
            }
        }

        public string Login_password
        {
            get
            {
                return login_password;
            }
            set
            {
                if (login_password != value)
                {
                    login_password = value;
                    NotifyOfPropertyChange(() => Login_password);
                }
            }
        }

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

        public decimal Post_no
        {
            get
            {
                return post_no;
            }
            set
            {
                if (post_no != value)
                {
                    post_no = value;
                    NotifyOfPropertyChange(() => Post_no);
                }
            }
        }

        public string Emsize_last_name
        {
            get
            {
                return emsize_last_name;
            }
            set
            {
                if (emsize_last_name != value)
                {
                    emsize_last_name = value;
                    NotifyOfPropertyChange(() => Emsize_last_name);
                }
            }
        }

        public string Emsize_first_name
        {
            get
            {
                return emsize_first_name;
            }
            set
            {
                if (emsize_first_name != value)
                {
                    emsize_first_name = value;
                    NotifyOfPropertyChange(() => Emsize_first_name);
                }
            }
        }

        public string Alphabet_last_name
        {
            get
            {
                return alphabet_last_name;
            }
            set
            {
                if (alphabet_last_name != value)
                {
                    alphabet_last_name = value;
                    NotifyOfPropertyChange(() => Alphabet_last_name);
                }
            }
        }

        public string Alphabet_first_name
        {
            get
            {
                return alphabet_first_name;
            }
            set
            {
                if (alphabet_first_name != value)
                {
                    alphabet_first_name = value;
                    NotifyOfPropertyChange(() => Alphabet_first_name);
                }
            }
        }

        public decimal? Time_table_no
        {
            get
            {
                return time_table_no;
            }
            set
            {
                if (time_table_no != value)
                {
                    time_table_no = value;
                    NotifyOfPropertyChange(() => Time_table_no);
                }
            }
        }

        public decimal Use_flag_of_holiday
        {
            get
            {
                return use_flag_of_holiday;
            }
            set
            {
                if (use_flag_of_holiday != value)
                {
                    use_flag_of_holiday = value;
                    NotifyOfPropertyChange(() => Use_flag_of_holiday);
                }
            }
        }

        public decimal? Etirement_date
        {
            get
            {
                return etirement_date;
            }
            set
            {
                if (etirement_date != value)
                {
                    etirement_date = value;
                    NotifyOfPropertyChange(() => Etirement_date);
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

        public decimal Status
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
        public string Post_name
        {
            get
            {
                return post_name;
            }
            set
            {
                if (post_name != value)
                {
                    post_name = value;
                    NotifyOfPropertyChange(() => Post_name);
                }
            }
        }


        #endregion
    }
}
