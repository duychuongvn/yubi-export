
namespace YUBI_TOOL.Model
{
    public class TerminalSettingModel : ModelBase
    {
        private decimal machine_no;
        private decimal use_type;
        private decimal device;
        private string ip_address;
        private decimal netport_no;
        private int status;
        public decimal Machine_no
        {
            get
            {
                return machine_no;
            }
            set
            {
                if (machine_no != value)
                {
                    machine_no = value;
                    NotifyOfPropertyChange(() => Machine_no);
                }
            }
        }

        public decimal Use_type
        {
            get
            {
                return use_type;
            }
            set
            {
                if (use_type != value)
                {
                    use_type = value;
                    NotifyOfPropertyChange(() => Use_type);
                }
            }
        }

        public decimal Device
        {
            get
            {
                return device;
            }
            set
            {
                if (device != value)
                {
                    device = value;
                    NotifyOfPropertyChange(() => Device);
                }
            }
        }

        public string Ip_address
        {
            get
            {
                return ip_address;
            }
            set
            {
                if (ip_address != value)
                {
                    ip_address = value;
                    NotifyOfPropertyChange(() => Ip_address);
                }
            }
        }

        public decimal Netport_no
        {
            get
            {
                return netport_no;
            }
            set
            {
                if (netport_no != value)
                {
                    netport_no = value;
                    NotifyOfPropertyChange(() => Netport_no);
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
