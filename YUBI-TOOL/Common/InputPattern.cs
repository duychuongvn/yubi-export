using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;

namespace YUBI_TOOL.Common
{
    public class InputPattern
    {
        private static DependencyProperty TextPatternProperty = DependencyProperty.RegisterAttached("TextPattern", typeof(string),
            typeof(InputPattern), new UIPropertyMetadata(null, TextPatternChanged));
        private static DependencyProperty UseDeleteKeyProperty = DependencyProperty.RegisterAttached("UseDeleteKey", typeof(bool),
            typeof(InputPattern), new UIPropertyMetadata(false, UseDeleteKeyChanged));

        private static void UseDeleteKeyChanged(object sender, DependencyPropertyChangedEventArgs e)
        {
            UIElement control = sender as UIElement;
            control.PreviewKeyUp -= new System.Windows.Input.KeyEventHandler(control_KeyUp);
            control.PreviewKeyUp += new System.Windows.Input.KeyEventHandler(control_KeyUp);
            control.PreviewKeyDown -= new System.Windows.Input.KeyEventHandler(control_KeyDown);
            control.PreviewKeyDown += new System.Windows.Input.KeyEventHandler(control_KeyDown);
        }
        private static void TextPatternChanged(object sender, DependencyPropertyChangedEventArgs e)
        {
            UIElement control = sender as UIElement;
            control.PreviewKeyUp -= new System.Windows.Input.KeyEventHandler(control_KeyUp);
            control.PreviewKeyUp += new System.Windows.Input.KeyEventHandler(control_KeyUp);
            control.PreviewKeyDown -= new System.Windows.Input.KeyEventHandler(control_KeyDown);
            control.PreviewKeyDown += new System.Windows.Input.KeyEventHandler(control_KeyDown);
        }

        private static void control_KeyDown(object sender, System.Windows.Input.KeyEventArgs e)
        {
            UIElement control = sender as UIElement;
            string textRegex = GetTextPattern(control);
            if (!string.IsNullOrEmpty(textRegex))
            {
                string input = e.Key.ToString();
                if (input.Length == 1)
                {
                    Match match = Regex.Match(input, textRegex);
                    if (!match.Success)
                    {
                        e.Handled = true;
                    }
                }
            }
            if (e.Key == System.Windows.Input.Key.Right || e.Key == System.Windows.Input.Key.Left
                || e.Key == System.Windows.Input.Key.Up || e.Key == System.Windows.Input.Key.Down)
            {
                DataGridCell cell = FindAncestor<DataGridCell>(control);
                if (cell != null)
                {
                    cell.IsEditing = false;
                }
            }
            if (control is TextBlock && e.Key == System.Windows.Input.Key.Delete)
            {
                DataGridCell cell = FindAncestor<DataGridCell>(control);
                if (cell != null)
                {
                    cell.IsEditing = true;
                }
            }
        }
        //EventManager.RegisterClassHandler(typeof(TabControl), TabControl.GotKeyboardFocusEvent, new KeyboardFocusChangedEventHandler(Event), true);
        private static void control_KeyUp(object sender, System.Windows.Input.KeyEventArgs e)
        {
            Control control = sender as Control;
            string textRegex = GetTextPattern(control);
            if (!string.IsNullOrEmpty(textRegex))
            {
                string input = e.Key.ToString();
                if (input.Length == 1)
                {
                    Match match = Regex.Match(input, textRegex);
                    if (!match.Success)
                    {
                        e.Handled = true;
                    }
                }
            }
        }
      
        public static T FindAncestor<T>(DependencyObject dependencyObject) where T : DependencyObject
        {
            if (dependencyObject is T)
            {
                return dependencyObject as T;
            }
            var parent = VisualTreeHelper.GetParent(dependencyObject);

            if (parent == null)
            {
                return null;
            }
            T parentT = parent as T;
            return parentT ?? FindAncestor<T>(parent);
        }
        public static void SetTextPattern(UIElement control, string value)
        {
            control.SetValue(TextPatternProperty, value);
        }
        public static string GetTextPattern(UIElement control)
        {

            return (string)control.GetValue(TextPatternProperty);
        }
        public static void SetUseDeleteKey(UIElement control, bool value)
        {
            control.SetValue(UseDeleteKeyProperty, value);
        }
        public static bool GetUseDeleteKey(UIElement control)
        {

            return (bool)control.GetValue(UseDeleteKeyProperty);
        }
    }
}
