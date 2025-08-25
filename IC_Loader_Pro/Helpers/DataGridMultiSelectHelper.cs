using System.Collections;
using System.Collections.Specialized;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Controls;
using static IC_Loader_Pro.Module1;

namespace IC_Loader_Pro.Helpers
{
    public static class DataGridSelectedItemsBehavior
    {
        private static readonly ConditionalWeakTable<object, DataGrid> _associations =
            new ConditionalWeakTable<object, DataGrid>();

        private static bool _isSyncing;

        public static readonly DependencyProperty SelectedItemsProperty =
            DependencyProperty.RegisterAttached("SelectedItems", typeof(IList), typeof(DataGridSelectedItemsBehavior), new PropertyMetadata(null, OnSelectedItemsChanged));

        public static IList GetSelectedItems(DependencyObject d) => (IList)d.GetValue(SelectedItemsProperty);
        public static void SetSelectedItems(DependencyObject d, IList value) => d.SetValue(SelectedItemsProperty, value);

        private static void OnSelectedItemsChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            if (!(d is DataGrid dataGrid)) return;

            if (e.OldValue != null)
            {
                _associations.Remove(e.OldValue);
                if (e.OldValue is INotifyCollectionChanged oldCollection)
                {
                    oldCollection.CollectionChanged -= OnViewModelCollectionChanged;
                }
            }
            dataGrid.SelectionChanged -= OnGridSelectionChanged;

            if (e.NewValue is INotifyCollectionChanged newCollection)
            {
                _associations.Add(newCollection, dataGrid);
                newCollection.CollectionChanged += OnViewModelCollectionChanged;
                dataGrid.SelectionChanged += OnGridSelectionChanged;
                SyncToDataGrid(dataGrid, newCollection as IList);
            }
        }

        private static void OnViewModelCollectionChanged(object sender, NotifyCollectionChangedEventArgs e)
        {
            //Log.RecordMessage($"DataGrid Behavior: ViewModel's collection changed (Action: {e.Action}).", BIS_Log.BisLogMessageType.Note);

            // ** THIS IS THE NEW DIAGNOSTIC LOGIC **
            if (_associations.TryGetValue(sender, out DataGrid dataGrid))
            {
            //    Log.RecordMessage("--> Association FOUND for the collection. Proceeding to sync grid.", BIS_Log.BisLogMessageType.Note);
                SyncToDataGrid(dataGrid, sender as IList);
            }
            else
            {
               // Log.RecordMessage("--> Association NOT FOUND for the collection. Cannot sync grid.", BIS_Log.BisLogMessageType.Note);
            }
        }

        private static void OnGridSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (_isSyncing) return;
            if (sender is DataGrid dataGrid)
            {
                var viewModelCollection = GetSelectedItems(dataGrid);
                if (viewModelCollection == null) return;

                _isSyncing = true;
                viewModelCollection.Clear();
                foreach (var item in dataGrid.SelectedItems)
                {
                    viewModelCollection.Add(item);
                }
                _isSyncing = false;
            }
        }

        private static void SyncToDataGrid(DataGrid dataGrid, IList viewModelCollection)
        {
            if (_isSyncing) return;

            _isSyncing = true;

            dataGrid.SelectedItems.Clear();
            if (viewModelCollection != null)
            {
                foreach (var item in viewModelCollection)
                {
                    dataGrid.SelectedItems.Add(item);
                }
            }

            _isSyncing = false;

            if (viewModelCollection != null && viewModelCollection.Count > 0)
            {
                dataGrid.Focus();
                dataGrid.ScrollIntoView(viewModelCollection[0]);
            }
        }
    }
}