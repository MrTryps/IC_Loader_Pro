using IC_Loader_Pro.Models;
using System.Collections.ObjectModel;
using System.Windows.Data;

namespace IC_Loader_Pro
{
    /// <summary>
    /// This ViewModel is used only by the Visual Studio XAML designer
    /// to provide sample data and avoid running live code.
    /// </summary>
    internal class DesignTimeViewModel : Dockpane_IC_LoaderViewModel
    {
        private readonly object _lockQueueCollection = new object();

        public DesignTimeViewModel()
        {
            // Create some sample queue summaries for the designer
            var sampleQueues = new ObservableCollection<ICQueueSummary>
            {
                new ICQueueSummary { Name = "CEAs", EmailCount = 12 },
                new ICQueueSummary { Name = "DNAs", EmailCount = 5 },
                new ICQueueSummary { Name = "WRAs", EmailCount = 21 }
            };

            // Use the public property from the base class to set the design-time data
            _ListOfIcEmailTypeSummaries = sampleQueues;
            PublicListOfIcEmailTypeSummaries = new ReadOnlyObservableCollection<ICQueueSummary>(_ListOfIcEmailTypeSummaries);

            // Enable collection synchronization for the designer
            BindingOperations.EnableCollectionSynchronization(PublicListOfIcEmailTypeSummaries, _lockQueueCollection);

            // Set a default selected item for the designer
            SelectedIcType = PublicListOfIcEmailTypeSummaries[0];
            StatusMessage = "Design Time - Ready";
        }

        // We need to shadow the base properties to provide a public setter for the collection
        // This is a common pattern for design-time data contexts.
        public new ReadOnlyObservableCollection<ICQueueSummary> PublicListOfIcEmailTypeSummaries { get; }
        private new ObservableCollection<ICQueueSummary> _ListOfIcEmailTypeSummaries;
    }
}