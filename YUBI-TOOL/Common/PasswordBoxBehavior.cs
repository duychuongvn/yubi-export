using System.Windows;
using System.Windows.Controls;
using System.Windows.Interactivity;

namespace YUBI_TOOL.Common
{
    public class PasswordBoxBehavior : Behavior<PasswordBox>
    {
        private static DependencyProperty PasswordValueProperty = DependencyProperty.Register("PasswordValue", typeof(string), typeof(PasswordBoxBehavior));
        protected override void OnAttached()
        {
            base.OnAttached();

            AssociatedObject.PasswordChanged += new RoutedEventHandler(AssociatedObject_PasswordChanged);

        }

        protected override void OnDetaching()
        {
            base.OnDetaching();
            AssociatedObject.PasswordChanged -= new RoutedEventHandler(AssociatedObject_PasswordChanged);
        }
        private void AssociatedObject_PasswordChanged(object sender, RoutedEventArgs e)
        {
            PasswordValue = AssociatedObject.Password;
        }

        protected override void OnPropertyChanged(DependencyPropertyChangedEventArgs e)
        {
            base.OnPropertyChanged(e);
            if (e.Property == PasswordValueProperty)
            {
                if (PasswordValue != AssociatedObject.Password)
                {
                    AssociatedObject.Password = PasswordValue;
                }
            }

        }
        public string PasswordValue
        {
            get { return GetValue(PasswordValueProperty) as string; }
            set
            {
                SetValue(PasswordValueProperty, value);
            }
        }
    }
}
