using System.Windows.Controls;
using System.Windows.Interactivity;

namespace YUBI_TOOL.Common
{
    /// <summary>
    /// Custom behavior that allows for DataGrid Rows to be ReadOnly on per-row basis
    /// </summary>
    public class DataGridRowReadOnlyBehavior : Behavior<DataGrid>
    {
        protected override void OnAttached()
        {
            base.OnAttached();
            AssociatedObject.BeginningEdit += AssociatedObject_BeginningEdit;
        }

        private void AssociatedObject_BeginningEdit(object sender, DataGridBeginningEditEventArgs e)
        {
            var isReadOnlyRow = ReadOnlyService.GetIsReadOnly(e.Row);
            if (isReadOnlyRow)
            {
                e.Cancel = true;
            }
        }

        protected override void OnDetaching()
        {
            AssociatedObject.BeginningEdit -= AssociatedObject_BeginningEdit;
        }
    }
}
