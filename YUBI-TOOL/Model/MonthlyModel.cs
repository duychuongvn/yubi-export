
namespace YUBI_TOOL.Model
{
   public class MonthlyModel: ModelBase
    {
       private decimal company_no;
       private decimal? expiration_from;
       private decimal? expiration_to;
       private decimal? cutoff_day;
       private decimal unit_minutes;

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

       public decimal? Cutoff_day
       {
           get
           {
               return cutoff_day;
           }
           set
           {
               if (cutoff_day != value)
               {
                   cutoff_day = value;
                   NotifyOfPropertyChange(() => Cutoff_day);
               }
           }
       }

       public decimal Unit_minutes
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

    }
}
