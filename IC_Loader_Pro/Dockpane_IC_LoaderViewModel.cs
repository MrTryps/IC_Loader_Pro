using ArcGIS.Core.Geometry;
using ArcGIS.Desktop.Core;
using ArcGIS.Desktop.Core.Events;
using ArcGIS.Desktop.Core.Geoprocessing;
using ArcGIS.Desktop.Framework;
using ArcGIS.Desktop.Framework.Contracts;
using ArcGIS.Desktop.Framework.Threading.Tasks;
using ArcGIS.Desktop.Mapping;
using ArcGIS.Desktop.Mapping.Events;
using BIS_Tools_DataModels_2025;
using IC_Loader_Pro.Models; // Your ICQueueSummary  class
using System;
using System.Collections.Generic;
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
        private readonly ObservableCollection<ICQueueSummary > _ListOfIcEmailTypeSummaries = new ObservableCollection<ICQueueSummary >();
        // This is a read-only wrapper around the real list that we will expose to the UI
        private readonly ReadOnlyObservableCollection<ICQueueSummary > _readOnly_ListOfIcEmailTypeSummaries;

        private List<EmailItem> _emailsForCurrentQueue;
        private int _currentEmailIndex = -1;

        // This property holds the SINGLE email we are currently processing
        private EmailItem _currentEmail;
        public EmailItem CurrentEmail
        {
            get => _currentEmail;
            set => SetProperty(ref _currentEmail, value);
        }

        private ICQueueSummary  _selectedICEmailSummary;
        private bool _isInitialized = false;
        private readonly object _lock = new object();
        private string _statusMessage = "Please open or create an ArcGIS Pro project.";
        public string StatusMessage { get => _statusMessage; set => SetProperty(ref _statusMessage, value); }

        private IcGisTypeSetting IcGisTypeSetting { get; set; }


        #endregion

        #region Constructor
        protected Dockpane_IC_LoaderViewModel()
        {
           
            // Create the public, read-only collection that the UI will bind to
            _readOnly_ListOfIcEmailTypeSummaries = new ReadOnlyObservableCollection<ICQueueSummary >(_ListOfIcEmailTypeSummaries);

            // This is a key step from the sample. It allows a background thread to safely update a collection that the UI is bound to.
            BindingOperations.EnableCollectionSynchronization(_readOnly_ListOfIcEmailTypeSummaries, _lockQueueCollection);

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
     
        #region Public Properties for UI Binding

        /// <summary>
        /// The list of IC Queues exposed to the View.
        /// </summary>
        public ReadOnlyObservableCollection<ICQueueSummary > ICQueues => _readOnly_ListOfIcEmailTypeSummaries;

        /// <summary>
        /// The currently selected IC Queue from the UI.
        /// </summary>
        public ICQueueSummary  SelectedQueue
        {
            get => _selectedICEmailSummary;
            set
            {
                // SetProperty is a helper method from the DockPane base class
                SetProperty(ref _selectedICEmailSummary, value, () => SelectedQueue);
                // When a queue is selected, kick off the process to load its emails.
                if (_selectedICEmailSummary != null && IsUIEnabled)
                {
                    _ = LoadEmailsForQueueAsync(_selectedICEmailSummary);
                }
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
            ActiveMapViewChangedEvent.Subscribe(OnActiveMapViewChanged);
            //ProjectOpenedEvent.Subscribe(OnProjectOpened);
            //ProjectClosedEvent.Subscribe(OnProjectClosed);

            //// This handles the case where the dockpane is opened AFTER a project is already open.
            //if (Project.Current != null)
            //{
            //    OnProjectOpened(new ProjectEventArgs(Project.Current));
            //}
            //return base.InitializeAsync();
            if (MapView.Active != null)
            {
                OnActiveMapViewChanged(new ActiveMapViewChangedEventArgs(MapView.Active, null));
            }
            return Task.CompletedTask;
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

        private void OnActiveMapViewChanged(ActiveMapViewChangedEventArgs args)
        {
            // The new, incoming view is in the 'IncomingView' property.
            // If it's null, it means no map view is active (e.g., all maps were closed).
            if (args.IncomingView == null)
            {
                IsUIEnabled = false;
                StatusMessage = "Please open a map view to begin.";
                return;
            }

            // Now that we know a map view is active, kick off our initialization.
            // We pass the Map from the incoming view.
            _ = LoadAndInitializeAsync(args.IncomingView.Map);
        }



        ///// <summary>
        ///// This is our main entry point, called only when a project is confirmed to be open.
        ///// </summary>
        //private void OnProjectOpened(ProjectEventArgs args)
        //{
        //    _ = LoadAndInitializeAsync();
        //}

        ///// <summary>
        ///// Resets the UI when the user closes their project.
        ///// </summary>
        //private void OnProjectClosed(ProjectEventArgs args)
        //{
        //    lock (_lock) { _isInitialized = false; }
        //    IsUIEnabled = false;
        //    StatusMessage = "Please open or create an ArcGIS Pro project.";
        //}

        #endregion

    }
}