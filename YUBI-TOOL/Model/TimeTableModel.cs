
namespace YUBI_TOOL.Model
{
    public class TimeTableModel : ModelBase
    {
        private decimal time_table_no;
        private decimal? expiration_from;
        private decimal? expiration_to;
        private string time_table_name;
        private string abbreviation;
        private decimal work_from;
        private decimal work_to;
        private decimal? coretime_from;
        private decimal delimitation;
        private decimal? coretime_to;
        private decimal? midnight_work_from;
        private decimal? midnight_work_to;
        private decimal? rest1_from;
        private decimal? rest1_to;
        private decimal? rest10_from;
        private decimal? rest10_to;
        private decimal? rest2_from;
        private decimal? rest2_to;
        private decimal? rest3_from;
        private decimal? rest3_to;
        private decimal? rest4_from;
        private decimal? rest4_to;
        private decimal? rest5_from;
        private decimal? rest5_to;
        private decimal? rest6_from;
        private decimal? rest6_to;
        private decimal? rest7_from;
        private decimal? rest7_to;
        private decimal? rest8_from;
        private decimal? rest8_to;
        private decimal? rest9_from;
        private decimal? rest9_to;
        private decimal? unit_minutes;
        private decimal? over_unit_minutes;
        private decimal status;
        public decimal Time_table_no
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

        public string Time_table_name
        {
            get
            {
                return time_table_name;
            }
            set
            {
                if (time_table_name != value)
                {
                    time_table_name = value;
                    NotifyOfPropertyChange(() => Time_table_name);
                }
            }
        }

        public string Abbreviation
        {
            get
            {
                return abbreviation;
            }
            set
            {
                if (abbreviation != value)
                {
                    abbreviation = value;
                    NotifyOfPropertyChange(() => Abbreviation);
                }
            }
        }

        public decimal Work_from
        {
            get
            {
                return work_from;
            }
            set
            {
                if (work_from != value)
                {
                    work_from = value;
                    NotifyOfPropertyChange(() => Work_from);
                }
            }
        }

        public decimal Work_to
        {
            get
            {
                return work_to;
            }
            set
            {
                if (work_to != value)
                {
                    work_to = value;
                    NotifyOfPropertyChange(() => Work_to);
                }
            }
        }

        public decimal? Coretime_from
        {
            get
            {
                return coretime_from;
            }
            set
            {
                if (coretime_from != value)
                {
                    coretime_from = value;
                    NotifyOfPropertyChange(() => Coretime_from);
                }
            }
        }

        public decimal Delimitation
        {
            get
            {
                return delimitation;
            }
            set
            {
                if (delimitation != value)
                {
                    delimitation = value;
                    NotifyOfPropertyChange(() => Delimitation);
                }
            }
        }

        public decimal? Coretime_to
        {
            get
            {
                return coretime_to;
            }
            set
            {
                if (coretime_to != value)
                {
                    coretime_to = value;
                    NotifyOfPropertyChange(() => Coretime_to);
                }
            }
        }

        public decimal? Midnight_work_from
        {
            get
            {
                return midnight_work_from;
            }
            set
            {
                if (midnight_work_from != value)
                {
                    midnight_work_from = value;
                    NotifyOfPropertyChange(() => Midnight_work_from);
                }
            }
        }

        public decimal? Midnight_work_to
        {
            get
            {
                return midnight_work_to;
            }
            set
            {
                if (midnight_work_to != value)
                {
                    midnight_work_to = value;
                    NotifyOfPropertyChange(() => Midnight_work_to);
                }
            }
        }

        public decimal? Rest1_from
        {
            get
            {
                return rest1_from;
            }
            set
            {
                if (rest1_from != value)
                {
                    rest1_from = value;
                    NotifyOfPropertyChange(() => Rest1_from);
                }
            }
        }

        public decimal? Rest1_to
        {
            get
            {
                return rest1_to;
            }
            set
            {
                if (rest1_to != value)
                {
                    rest1_to = value;
                    NotifyOfPropertyChange(() => Rest1_to);
                }
            }
        }
        public decimal? Rest10_from
        {
            get
            {
                return rest10_from;
            }
            set
            {
                if (rest10_from != value)
                {
                    rest10_from = value;
                    NotifyOfPropertyChange(() => Rest10_from);
                }
            }
        }

        public decimal? Rest10_to
        {
            get
            {
                return rest10_to;
            }
            set
            {
                if (rest10_to != value)
                {
                    rest10_to = value;
                    NotifyOfPropertyChange(() => Rest10_to);
                }
            }
        }
        public decimal? Rest2_from
        {
            get
            {
                return rest2_from;
            }
            set
            {
                if (rest2_from != value)
                {
                    rest2_from = value;
                    NotifyOfPropertyChange(() => Rest2_from);
                }
            }
        }

        public decimal? Rest2_to
        {
            get
            {
                return rest2_to;
            }
            set
            {
                if (rest2_to != value)
                {
                    rest2_to = value;
                    NotifyOfPropertyChange(() => Rest2_to);
                }
            }
        }

        public decimal? Rest3_from
        {
            get
            {
                return rest3_from;
            }
            set
            {
                if (rest3_from != value)
                {
                    rest3_from = value;
                    NotifyOfPropertyChange(() => Rest3_from);
                }
            }
        }

        public decimal? Rest3_to
        {
            get
            {
                return rest3_to;
            }
            set
            {
                if (rest3_to != value)
                {
                    rest3_to = value;
                    NotifyOfPropertyChange(() => Rest3_to);
                }
            }
        }

        public decimal? Rest4_from
        {
            get
            {
                return rest4_from;
            }
            set
            {
                if (rest4_from != value)
                {
                    rest4_from = value;
                    NotifyOfPropertyChange(() => Rest4_from);
                }
            }
        }

        public decimal? Rest4_to
        {
            get
            {
                return rest4_to;
            }
            set
            {
                if (rest4_to != value)
                {
                    rest4_to = value;
                    NotifyOfPropertyChange(() => Rest4_to);
                }
            }
        }

        public decimal? Rest5_from
        {
            get
            {
                return rest5_from;
            }
            set
            {
                if (rest5_from != value)
                {
                    rest5_from = value;
                    NotifyOfPropertyChange(() => Rest5_from);
                }
            }
        }

        public decimal? Rest5_to
        {
            get
            {
                return rest5_to;
            }
            set
            {
                if (rest5_to != value)
                {
                    rest5_to = value;
                    NotifyOfPropertyChange(() => Rest5_to);
                }
            }
        }

        public decimal? Rest6_from
        {
            get
            {
                return rest6_from;
            }
            set
            {
                if (rest6_from != value)
                {
                    rest6_from = value;
                    NotifyOfPropertyChange(() => Rest6_from);
                }
            }
        }

        public decimal? Rest6_to
        {
            get
            {
                return rest6_to;
            }
            set
            {
                if (rest6_to != value)
                {
                    rest6_to = value;
                    NotifyOfPropertyChange(() => Rest6_to);
                }
            }
        }

        public decimal? Rest7_from
        {
            get
            {
                return rest7_from;
            }
            set
            {
                if (rest7_from != value)
                {
                    rest7_from = value;
                    NotifyOfPropertyChange(() => Rest7_from);
                }
            }
        }

        public decimal? Rest7_to
        {
            get
            {
                return rest7_to;
            }
            set
            {
                if (rest7_to != value)
                {
                    rest7_to = value;
                    NotifyOfPropertyChange(() => Rest7_to);
                }
            }
        }

        public decimal? Rest8_from
        {
            get
            {
                return rest8_from;
            }
            set
            {
                if (rest8_from != value)
                {
                    rest8_from = value;
                    NotifyOfPropertyChange(() => Rest8_from);
                }
            }
        }

        public decimal? Rest8_to
        {
            get
            {
                return rest8_to;
            }
            set
            {
                if (rest8_to != value)
                {
                    rest8_to = value;
                    NotifyOfPropertyChange(() => Rest8_to);
                }
            }
        }

        public decimal? Rest9_from
        {
            get
            {
                return rest9_from;
            }
            set
            {
                if (rest9_from != value)
                {
                    rest9_from = value;
                    NotifyOfPropertyChange(() => Rest9_from);
                }
            }
        }

        public decimal? Rest9_to
        {
            get
            {
                return rest9_to;
            }
            set
            {
                if (rest9_to != value)
                {
                    rest9_to = value;
                    NotifyOfPropertyChange(() => Rest9_to);
                }
            }
        }

        public decimal? Unit_minutes
        {
            get
            {
                return unit_minutes;
            }
            set
            {
                if (unit_minutes != value)
                {
                    unit_minutes = value;
                    NotifyOfPropertyChange(() => Unit_minutes);
                }
            }
        }

        public decimal? Over_unit_minutes
        {
            get
            {
                return over_unit_minutes;
            }
            set
            {
                if (over_unit_minutes != value)
                {
                    over_unit_minutes = value;
                    NotifyOfPropertyChange(() => Over_unit_minutes);
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
