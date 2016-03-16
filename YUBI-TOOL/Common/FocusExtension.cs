using System.Windows;
using System.Windows.Input;
using System.Windows.Threading;

namespace YUBI_TOOL.Common
{
    public class FocusExtension
    {
        private const string PROP_IS_FOCUSED = "IsFocused";
        public static readonly DependencyProperty IsFocusedProperty =
            DependencyProperty.RegisterAttached(PROP_IS_FOCUSED, typeof(bool), typeof(FocusExtension), new UIPropertyMetadata(false, OnIsFocusedPropertyChanged));

        private static void OnIsFocusedPropertyChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            var uie = (UIElement)d;

            if ((bool)e.NewValue)
            {
                
                    uie.Dispatcher.BeginInvoke(new System.Action(delegate
                    {
                        uie.Focus();
                        Keyboard.Focus(uie);
                        SetIsFocused(d, false);
                    }), DispatcherPriority.Background, null);
                }
            
        }

        public static bool GetIsFocused(DependencyObject obj)
        {
            return (bool)obj.GetValue(IsFocusedProperty);
        }

        public static void SetIsFocused(DependencyObject obj, bool value)
        {
            obj.SetValue(IsFocusedProperty, value);
        }
    }
}
