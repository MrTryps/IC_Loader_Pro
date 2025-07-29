using System.Collections;
using System.Windows;
using System.Windows.Controls;

namespace IC_Loader_Pro.Helpers
{
    public class DataGridSelectedItemsBehavior
    {
        public static readonly DependencyProperty SelectedItemsProperty =
            DependencyProperty.RegisterAttached("SelectedItems", typeof(IList), typeof(DataGridSelectedItemsBehavior), new PropertyMetadata(null, OnSelectedItemsChanged));

        public static IList GetSelectedItems(DependencyObject d)
        {
            return (IList)d.GetValue(SelectedItemsProperty);
        }

        public static void SetSelectedItems(DependencyObject d, IList value)
        {
            d.SetValue(SelectedItemsProperty, value);
        }

        private static void OnSelectedItemsChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            if (d is DataGrid grid)
            {
                grid.SelectionChanged -= OnGridSelectionChanged;
                if (e.NewValue != null)
                {
                    grid.SelectionChanged += OnGridSelectionChanged;
                }
            }
        }

        private static void OnGridSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (sender is DataGrid grid)
            {
                IList selectedItems = GetSelectedItems(grid);
                if (selectedItems == null) return;

                selectedItems.Clear();
                foreach (var item in grid.SelectedItems)
                {
                    selectedItems.Add(item);
                }
            }
        }
    }
}