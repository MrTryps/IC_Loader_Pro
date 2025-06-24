using ArcGIS.Core.Geometry;
using ArcGIS.Desktop.Core;
using ArcGIS.Desktop.Core.Events;
using ArcGIS.Desktop.Core.Geoprocessing;
using ArcGIS.Desktop.Framework;
using ArcGIS.Desktop.Framework.Contracts;
using ArcGIS.Desktop.Framework.Threading.Tasks;
using ArcGIS.Desktop.Mapping;
using IC_Loader_Pro.Models; // Your ICQueueSummary class
using System;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Data;
using System.Windows.Input;
using static BIS_Tools_2025_Core.BIS_Log;
using static IC_Loader_Pro.Module1; // For  Log

namespace IC_Loader_Pro
{
    internal partial class  Dockpane_IC_LoaderViewModel : DockPane
    {
        #region Private Members
        /// <summary>
        /// The unique ID of this dockpane, must match the ID in Config.daml
        /// </summary>
        public const string DockPaneId = "IC_Loader_Pro_Dockpane_IC_Loader";

        private readonly object _lockQueueCollection = new object();
        // This is the "real" list that we will add/remove items from
        private readonly ObservableCollection<ICQueueSummary> _ListOfIcEmailTypeSummaries = new ObservableCollection<ICQueueSummary>();
        // This is a read-only wrapper around the real list that we will expose to the UI
        private readonly ReadOnlyObservableCollection<ICQueueSummary> _readOnlyListOfQueues;

        private ICQueueSummary _selectedQueue;
        private bool _isInitialized = false;
        private readonly object _lock = new object();
        private string _statusMessage = "Please open or create an ArcGIS Pro project.";
        public string StatusMessage { get => _statusMessage; set => SetProperty(ref _statusMessage, value); }

        #endregion

        #region Constructor
        protected Dockpane_IC_LoaderViewModel()
        {
           
            // Create the public, read-only collection that the UI will bind to
            _readOnlyListOfQueues = new ReadOnlyObservableCollection<ICQueueSummary>(_ListOfIcEmailTypeSummaries);

            // This is a key step from the sample. It allows a background thread to safely update a collection that the UI is bound to.
            BindingOperations.EnableCollectionSynchronization(_readOnlyListOfQueues, _lockQueueCollection);

            // Initialize commands
            RefreshQueuesCommand = new RelayCommand(async () => await RefreshICQueuesAsync(), () => true);
            SaveCommand = new RelayCommand(async () => await OnSave(), () => IsUIEnabled);
            SkipCommand = new RelayCommand(async () => await OnSkip(), () => IsUIEnabled);
            RejectCommand = new RelayCommand(async () => await OnReject(), () => IsUIEnabled);
            ShowNotesCommand = new RelayCommand(async () => await OnShowNotes(), () => IsUIEnabled);
            SearchCommand = new RelayCommand(async () => await OnSearch(), () => IsUIEnabled);
            ToolsCommand = new RelayCommand(async () => await OnTools(), () => IsUIEnabled);
            OptionsCommand = new RelayCommand(async () => await OnOptions(), () => IsUIEnabled);
        }
        #endregion
     
        #region Public Properties and Commands for UI Binding

        /// <summary>
        /// The list of IC Queues exposed to the View.
        /// </summary>
        public ReadOnlyObservableCollection<ICQueueSummary> PublicListOfIcEmailTypeSummaries => _readOnlyListOfQueues;

        /// <summary>
        /// The currently selected IC Queue from the UI.
        /// </summary>
        public ICQueueSummary SelectedIcType
        {
            get => _selectedQueue;
            set
            {
                // SetProperty is a helper method from the DockPane base class
                SetProperty(ref _selectedQueue, value, () => SelectedIcType);
                // When a queue is selected, we can trigger logic here later
            }
        }

         #endregion

      

        #region Overrides and Static Show Method

        /// <summary>
        /// This is called by the framework when the dockpane is first created.
        /// Perfect for one-time setup like subscribing to events or loading initial data.
        /// </summary>
        protected override Task InitializeAsync()
        {
            ProjectOpenedEvent.Subscribe(OnProjectOpened);
            ProjectClosedEvent.Subscribe(OnProjectClosed);

            // This handles the case where the dockpane is opened AFTER a project is already open.
            if (Project.Current != null)
            {
                OnProjectOpened(new ProjectEventArgs(Project.Current));
            }
            return base.InitializeAsync();
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
        #region Event Handlers

        /// <summary>
        /// This is our main entry point, called only when a project is confirmed to be open.
        /// </summary>
        private void OnProjectOpened(ProjectEventArgs args)
        {
            _ = LoadAndInitializeAsync();
        }

        /// <summary>
        /// Resets the UI when the user closes their project.
        /// </summary>
        private void OnProjectClosed(ProjectEventArgs args)
        {
            lock (_lock) { _isInitialized = false; }
            IsUIEnabled = false;
            StatusMessage = "Please open or create an ArcGIS Pro project.";
        }

        #endregion

    }
}