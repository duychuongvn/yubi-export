
namespace YUBI_TOOL.Model
{
    public class PostModel : ModelBase
    {
        private string company_no;
        private string post_no;
        private decimal? expiration_from;
        private decimal? expiration_to;
        private string post_name;
        private decimal status;
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

        public string Post_no
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
    }
}
