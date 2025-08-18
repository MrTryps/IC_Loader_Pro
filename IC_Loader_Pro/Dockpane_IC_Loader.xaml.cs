using IC_Loader_Pro.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;


namespace IC_Loader_Pro
{
    /// <summary>
    /// Interaction logic for Dockpane_IC_LoaderView.xaml
    /// </summary>
    public partial class Dockpane_IC_LoaderView : UserControl
    {
        private Dockpane_IC_LoaderViewModel ViewModel => DataContext as Dockpane_IC_LoaderViewModel;

        public Dockpane_IC_LoaderView()
        {
            InitializeComponent();
        }

        private void DataGrid_GotFocus(object sender, RoutedEventArgs e)
        {
            if (ViewModel == null) return;

            if (sender is DataGrid focusedGrid)
            {
                // If the "Review" grid got focus, clear the selection in the "Use" grid.
                if (focusedGrid.ItemsSource == ViewModel.ShapesToReview)
                {
                    ViewModel.SelectedShapesToUse.Clear();
                }
                // If the "Use" grid got focus, clear the selection in the "Review" grid.
                else if (focusedGrid.ItemsSource == ViewModel.SelectedShapes)
                {
                    ViewModel.SelectedShapesForReview.Clear();
                }
            }
        }
       
    }
}
//System.Windows.Markup.XamlParseException: ''Provide value on 'System.Windows.Baml2006.TypeConverterMarkupExtension' threw an exception.' Line number '47' and line position '43'.'
