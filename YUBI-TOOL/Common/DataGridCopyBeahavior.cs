using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Interactivity;

namespace YUBI_TOOL.Common
{
    public class DataGridCopyBeahavior : Behavior<DataGrid>
    {
        public static readonly DependencyProperty CopyItemProperty =
           DependencyProperty.RegisterAttached("CopyItem", typeof(object), typeof(DataGridCopyBeahavior));
        public static readonly DependencyProperty PasteItemProperty =
        DependencyProperty.RegisterAttached("PasteItem", typeof(object), typeof(DataGridCopyBeahavior));
        public static readonly DependencyProperty SelectedDataItemProperty =
       DependencyProperty.RegisterAttached("SelectedDataItem", typeof(object), typeof(DataGridCopyBeahavior));
        public static readonly DependencyProperty DeletePathProperty =
       DependencyProperty.RegisterAttached("DeletePath", typeof(string), typeof(DataGridCopyBeahavior));

        protected override void OnAttached()
        {
            base.OnAttached();
            this.AssociatedObject.PreviewKeyDown += new System.Windows.Input.KeyEventHandler(AssociatedObject_PreviewKeyDown);
        }

        void AssociatedObject_PreviewKeyDown(object sender, System.Windows.Input.KeyEventArgs e)
        {
            if (e.Key == System.Windows.Input.Key.C && (Keyboard.Modifiers & ModifierKeys.Control) == ModifierKeys.Control)
            {
                e.Handled = true;
                PasteItem = null;
                CopyItem = AssociatedObject.CurrentCell.Item;
                
            }
            else if (e.Key == System.Windows.Input.Key.V && (Keyboard.Modifiers & ModifierKeys.Control) == ModifierKeys.Control)
            {
                e.Handled = true;
                PasteItem = AssociatedObject.CurrentCell.Item;
            }
            else if (e.Key == Key.Delete && !AssociatedObject.CurrentColumn.IsReadOnly)
            {
                SelectedDataItem = AssociatedObject.CurrentCell.Item;
                DeletePath = AssociatedObject.CurrentColumn.SortMemberPath;
            }
        }

        protected override void OnDetaching()
        {
            base.OnDetaching();
            this.AssociatedObject.PreviewKeyDown -= new System.Windows.Input.KeyEventHandler(AssociatedObject_PreviewKeyDown);
        }

        public object CopyItem
        {
            get { return GetValue(CopyItemProperty); }
            set { SetValue(CopyItemProperty, value); }
        }
        public object PasteItem
        {
            get { return GetValue(PasteItemProperty); }
            set { SetValue(PasteItemProperty, value); }
        }
        public object SelectedDataItem
        {
            get { return GetValue(SelectedDataItemProperty); }
            set { SetValue(SelectedDataItemProperty, value); }
        }
        public string DeletePath
        {
            get { return (string)GetValue(DeletePathProperty); }
            set { SetValue(DeletePathProperty, value); }
        }

    }
}
