using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Interactivity;
using System.Windows.Media;

namespace YUBI_TOOL.Common
{

    public class InputBehavior : Behavior<Control>
    {
        public string TextRegex { get; set; }
        protected override void OnAttached()
        {
            base.OnAttached();

            AssociatedObject.KeyUp += new System.Windows.Input.KeyEventHandler(AssociatedObject_KeyUp);
            AssociatedObject.KeyDown += new System.Windows.Input.KeyEventHandler(AssociatedObject_KeyDown);
        }

        protected override void OnDetaching()
        {
            base.OnDetaching();
            AssociatedObject.KeyUp -= new System.Windows.Input.KeyEventHandler(AssociatedObject_KeyUp);
            AssociatedObject.KeyDown -= new System.Windows.Input.KeyEventHandler(AssociatedObject_KeyDown);
        }
        private void AssociatedObject_KeyDown(object sender, System.Windows.Input.KeyEventArgs e)
        {
            if (!string.IsNullOrEmpty(TextRegex))
            {
                string input = e.Key.ToString();
                if (input.Length == 1)
                {
                    Match match = Regex.Match(input, TextRegex);
                    if (!match.Success)
                    {
                        e.Handled = true;
                    }
                }
            }
          
        }

        private void AssociatedObject_KeyUp(object sender, System.Windows.Input.KeyEventArgs e)
        {
            if (!string.IsNullOrEmpty(TextRegex))
            {
                string input = e.Key.ToString();
                if (input.Length == 1)
                {
                    Match match = Regex.Match(input, TextRegex);
                    if (!match.Success)
                    {
                        e.Handled = true;
                    }
                }
            }
            else if (sender is DataGrid)
            {
                if (e.Key == System.Windows.Input.Key.Tab)
                {
                    //DataGrid grid = sender as DataGrid;
                    //DataGridCell cell = FindAncestor<DataGridCell>(e.OriginalSource as UIElement);
                    //if (cell != null && cell.IsEditing)
                    //{
                    //    grid.CommitEdit(DataGridEditingUnit.Cell, true);
                    //    cell.MoveFocus(new TraversalRequest(FocusNavigationDirection.Next));
                    //}

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

    }
}
