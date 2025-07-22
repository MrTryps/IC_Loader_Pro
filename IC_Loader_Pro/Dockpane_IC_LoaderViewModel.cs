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
using IC_Loader_Pro.Models; // Your ICQueueSummary class
using IC_Loader_Pro.ViewModels;
using IC_Rules_2025;
using System;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Data;
using System.Windows.Input;
using static BIS_Log;
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

        // This collection will hold the filesets for the currently active email
        private readonly ObservableCollection<FileSetViewModel> _foundFileSets = new ObservableCollection<FileSetViewModel>();
        public ReadOnlyObservableCollection<FileSetViewModel> _readOnlyFoundFileSets { get; }


        private ICQueueSummary _selectedQueue;
        private bool _isInitialized = false;
        private readonly object _lock = new object();
        private string _statusMessage = "Please open or create an ArcGIS Pro project.";
        public string StatusMessage { get => _statusMessage; set => SetProperty(ref _statusMessage, value); }

        private string _currentEmailId;
        public string CurrentEmailId
        {
            get => _currentEmailId;
            set => SetProperty(ref _currentEmailId, value);
        }

        private IcTestResult _currentEmailTestResult;
        private AttachmentAnalysisResult _currentAttachmentAnalysis;

        private string _currentEmailSubject = "No email selected";
        public string CurrentEmailSubject
        {
            get => _currentEmailSubject;
            set => SetProperty(ref _currentEmailSubject, value);
        }

        private string _currentPrefId;
        public string CurrentPrefId
        {
            get => _currentPrefId;
            set => SetProperty(ref _currentPrefId, value);
        }

        private string _currentAltId;
        public string CurrentAltId
        {
            get => _currentAltId;
            set => SetProperty(ref _currentAltId, value);
        }

        private string _currentActivityNum;
        public string CurrentActivityNum
        {
            get => _currentActivityNum;
            set => SetProperty(ref _currentActivityNum, value);
        }

        private string _currentDelId;
        public string CurrentDelId
        {
            get => _currentDelId;
            set => SetProperty(ref _currentDelId, value);
        }

        #endregion

        #region Constructor
        protected Dockpane_IC_LoaderViewModel()
        {
           
            // Create the public, read-only collection that the UI will bind to
            _readOnlyListOfQueues = new ReadOnlyObservableCollection<ICQueueSummary>(_ListOfIcEmailTypeSummaries);
            _readOnlyFoundFileSets = new ReadOnlyObservableCollection<FileSetViewModel>(_foundFileSets);

            // This is a key step from the sample. It allows a background thread to safely update a collection that the UI is bound to.
            BindingOperations.EnableCollectionSynchronization(_readOnlyListOfQueues, _lockQueueCollection);

            // Initialize commands
            RefreshQueuesCommand = new RelayCommand(async () => await RefreshICQueuesAsync(), () => IsUIEnabled);
            SaveCommand = new RelayCommand(async () => await OnSave(), () => IsEmailActionEnabled);
            SkipCommand = new RelayCommand(async () => await OnSkip(), () => IsEmailActionEnabled);
            RejectCommand = new RelayCommand(async () => await OnReject(), () => IsEmailActionEnabled);
            ShowNotesCommand = new RelayCommand(async () => await OnShowNotes(), () => IsUIEnabled);
            SearchCommand = new RelayCommand(async () => await OnSearch(), () => IsUIEnabled);
            ToolsCommand = new RelayCommand(async () => await OnTools(), () => IsUIEnabled);
            OptionsCommand = new RelayCommand(async () => await OnOptions(), () => IsUIEnabled);
            ShowResultsCommand = new RelayCommand(OnShowResults, () => _currentEmailTestResult != null);
        }
        #endregion
     
        #region Public Properties and Commands for UI Binding

        /// <summary>
        /// The list of IC Queues exposed to the View.
        /// </summary>
        public ReadOnlyObservableCollection<ICQueueSummary> PublicListOfIcEmailTypeSummaries => _readOnlyListOfQueues;

        /// The list of identified filesets from the current email.
        /// </summary>
        public ReadOnlyObservableCollection<FileSetViewModel> FoundFileSets => _readOnlyFoundFileSets;

        /// <summary>
        /// The currently selected IC Queue from the UI.
        /// </summary>
        public ICQueueSummary SelectedIcType
        {
            get => _selectedQueue;
            set
            {
                if (_selectedQueue == value) return; // Don't re-process if the same queue is clicked again

                SetProperty(ref _selectedQueue, value);

                // If the new selection is not null, kick off the processing.
                // The underscore discards the returned Task, which is a standard
                // way to call an async method from a synchronous property setter.
                if (value != null)
                {
                    _ = ProcessSelectedQueueAsync();
                }


                // SetProperty is a helper method from the DockPane base class
               // SetProperty(ref _selectedQueue, value, () => SelectedIcType);
                // When a queue is selected, we can trigger logic here later
            }
        }

        #endregion

        private bool _isEmailActionEnabled = false;
        /// <summary>
        /// Controls whether the Save, Skip, and Reject buttons are enabled.
        /// </summary>
        public bool IsEmailActionEnabled
        {
            get => _isEmailActionEnabled;
            set => SetProperty(ref _isEmailActionEnabled, value);
        }


        #region Overrides and Static Show Method

        /// <summary>
        /// This is called by the framework when the dockpane is first created.
        /// Perfect for one-time setup like subscribing to events or loading initial data.
        /// </summary>
        protected override Task InitializeAsync()
        {
            ProjectOpenedEvent.Subscribe(OnProjectOpened);
            ProjectClosedEvent.Subscribe(OnProjectClosed);

            ActiveMapViewChangedEvent.Subscribe(OnActiveMapViewChanged);

            if (MapView.Active != null)
            {
                // A map is already active, so we can initialize immediately.
                OnActiveMapViewChanged(new ActiveMapViewChangedEventArgs(MapView.Active, null));
            }           
            return base.InitializeAsync();
        }

        private void OnActiveMapViewChanged(ActiveMapViewChangedEventArgs args)
        {
            if (MapView.Active != null && !_isInitialized)
            {
                // A map view is now active and we haven't run our setup yet.
                // This is the SAFE time to run your initialization.
                IcGisTypeSetting icSetting = null;
                _ = LoadAndInitializeAsync();

                // Set the flag so we don't run this full initialization again
                // every time the user switches between maps.
                _isInitialized = true;
            }
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
            Module1.Log.RecordMessage("Project opened. Waiting for active map view.", BisLogMessageType.Note );
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

        #region Private Helpers

        /// <summary>
        /// A simplified helper to force an action onto the ArcGIS Pro UI thread.
        /// </summary>
        private Task RunOnUIThread(Action action)
        {
            //Log.RecordMessage("Attempting to schedule action on UI thread...", BisLogMessageType.Note);
            if (OnUIThread)
            {
                // We are, so we can run the action directly.
                action();
                // Return a completed task.
                return Task.CompletedTask;
            }
            else
            {
                return Task.Factory.StartNew(action, CancellationToken.None, TaskCreationOptions.None, QueuedTask.UIScheduler);

            }
        }        

        /// <summary>
        /// Determines if the application is currently on the UI thread.
        /// </summary>
        private bool OnUIThread
        {
            get
            {
                if (FrameworkApplication.TestMode)
                    return QueuedTask.OnWorker;
                else
                    return System.Windows.Application.Current.Dispatcher.CheckAccess();
            }
        }
        #endregion


    }
}