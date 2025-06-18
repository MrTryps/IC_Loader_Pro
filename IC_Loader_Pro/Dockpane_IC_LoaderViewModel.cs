using ArcGIS.Desktop.Core;
using ArcGIS.Desktop.Core.Events;
using ArcGIS.Desktop.Framework;
using ArcGIS.Desktop.Framework.Contracts;
using ArcGIS.Desktop.Framework.Threading.Tasks;
using IC_Loader_Pro.Models; // Your ICQueueInfo class
using System.Collections.ObjectModel;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Data;
using System.Windows.Input;
using static IC_Loader_Pro.Module1; // For your Log

namespace IC_Loader_Pro
{
    internal class Dockpane_IC_LoaderViewModel : DockPane
    {
        #region Private Members
        /// <summary>
        /// The unique ID of this dockpane, must match the ID in Config.daml
        /// </summary>
        public const string DockPaneId = "IC_Loader_Pro_Dockpane_IC_Loader";

        private readonly object _lockQueueCollection = new object();

        // This is the "real" list that we will add/remove items from
        private readonly ObservableCollection<ICQueueInfo> _listOfQueues = new ObservableCollection<ICQueueInfo>();

        // This is a read-only wrapper around the real list that we will expose to the UI
        private readonly ReadOnlyObservableCollection<ICQueueInfo> _readOnlyListOfQueues;

        private ICQueueInfo _selectedQueue;
        #endregion

        #region Constructor
        protected Dockpane_IC_LoaderViewModel()
        {
            // Create the public, read-only collection that the UI will bind to
            _readOnlyListOfQueues = new ReadOnlyObservableCollection<ICQueueInfo>(_listOfQueues);

            // This is a key step from the sample. It allows a background thread to safely update a collection that the UI is bound to.
            BindingOperations.EnableCollectionSynchronization(_readOnlyListOfQueues, _lockQueueCollection);

            // Initialize commands
            RefreshQueuesCommand = new RelayCommand(async () => await RefreshICQueuesAsync(), () => true);
        }
        #endregion

        #region Public Properties and Commands for UI Binding

        /// <summary>
        /// The list of IC Queues exposed to the View.
        /// </summary>
        public ReadOnlyObservableCollection<ICQueueInfo> ICQueues => _readOnlyListOfQueues;

        /// <summary>
        /// The currently selected IC Queue from the UI.
        /// </summary>
        public ICQueueInfo SelectedQueue
        {
            get => _selectedQueue;
            set
            {
                // SetProperty is a helper method from the DockPane base class
                SetProperty(ref _selectedQueue, value, () => SelectedQueue);
                // When a queue is selected, we can trigger logic here later
            }
        }

        /// <summary>
        /// A command to trigger the refresh of the IC Queues.
        /// </summary>
        public ICommand RefreshQueuesCommand { get; private set; }

        #endregion

        #region Core Logic

        /// <summary>
        /// This method will contain the logic to call your Outlook library and get the real data.
        /// </summary>
        private Task RefreshICQueuesAsync()
        {
            return QueuedTask.Run(() =>
            {
                // This lock ensures that the collection is not modified by two threads at once.
                lock (_lockQueueCollection)
                {
                    _listOfQueues.Clear();

                    // For now, we use sample data. Later, we will replace this with a call
                    // to your BIS_IC_InputClasses_2025 library.
                    _listOfQueues.Add(new ICQueueInfo { Name = "Type A", EmailCount = 12 });
                    _listOfQueues.Add(new ICQueueInfo { Name = "Type B", EmailCount = 5 });
                    _listOfQueues.Add(new ICQueueInfo { Name = "Type C", EmailCount = 23 });
                }

                // Select the first item by default
                SelectedQueue = _readOnlyListOfQueues.FirstOrDefault();
            });
        }

        #endregion

        #region Overrides and Static Show Method

        /// <summary>
        /// This is called by the framework when the dockpane is first created.
        /// Perfect for one-time setup like subscribing to events or loading initial data.
        /// </summary>
        protected override Task InitializeAsync()
        {
            // We can add event subscriptions here if needed, like in the sample
            // ProjectOpenedEvent.Subscribe(...);

            // Let's load the queues when the dockpane initializes
            return RefreshICQueuesAsync();
        }

        /// <summary>
        /// The static method used by the button to show the dockpane.
        /// </summary>
        internal static void Show()
        {
            var pane = FrameworkApplication.DockPaneManager.Find(DockPaneId);
            if (pane == null)
                return;
            pane.Activate();
        }
        #endregion
    }
}