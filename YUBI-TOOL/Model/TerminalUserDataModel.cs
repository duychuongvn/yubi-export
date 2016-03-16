
namespace YUBI_TOOL.Model
{
    public class TerminalUserDataModel : ModelBase
    {
        private string id;
        private decimal backup_no;
        private decimal machine_privilege;
        private string enroll_data;
        public string Id
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

        public decimal Backup_no
        {
            get
            {
                return backup_no;
            }
            set
            {
                if (backup_no != value)
                {
                    backup_no = value;
                    NotifyOfPropertyChange(() => Backup_no);
                }
            }
        }

        public decimal Machine_privilege
        {
            get
            {
                return machine_privilege;
            }
            set
            {
                if (machine_privilege != value)
                {
                    machine_privilege = value;
                    NotifyOfPropertyChange(() => Machine_privilege);
                }
            }
        }

        public string Enroll_data
        {
            get
            {
                return enroll_data;
            }
            set
            {
                if (enroll_data != value)
                {
                    enroll_data = value;
                    NotifyOfPropertyChange(() => Enroll_data);
                }
            }
        }


    }
}
