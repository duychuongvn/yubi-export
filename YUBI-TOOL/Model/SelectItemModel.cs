
namespace YUBI_TOOL.Model
{
    public class SelectItemModel : ModelBase
    {
        private bool isSelected;
        private string itemCD;
        private string itemValue;
        public bool IsSelected
        {
            get
            {
                return isSelected;
            }
            set
            {
                if (isSelected != value)
                {
                    isSelected = value;
                    NotifyOfPropertyChange(() => IsSelected);
                }
            }
        }

        public string ItemCD
        {
            get
            {
                return itemCD;
            }
            set
            {
                if (itemCD != value)
                {
                    itemCD = value;
                    NotifyOfPropertyChange(() => ItemCD);
                }
            }
        }

        public string ItemValue
        {
            get
            {
                return itemValue;
            }
            set
            {
                if (itemValue != value)
                {
                    itemValue = value;
                    NotifyOfPropertyChange(() => ItemValue);
                }
            }
        }


    }
}
