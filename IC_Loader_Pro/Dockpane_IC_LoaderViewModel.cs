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
using IC_Loader_Pro.Models;
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


        #region Private Members & Collections

        private bool _isInitialized = false;
        private readonly object _lock = new object();
        private int _currentEmailIndex = -1;

        // Private modifiable collections
        private readonly object _lockQueueCollection = new object();
        private readonly ObservableCollection<ICQueueSummary> _ListOfIcEmailTypeSummaries = new ObservableCollection<ICQueueSummary>();

        private readonly object _lockShapesCollection = new object();
        private readonly ObservableCollection<ShapeItem> _listOfShapesToReview = new ObservableCollection<ShapeItem>();

        private readonly object _lockSelectedShapesCollection = new object();
        private readonly ObservableCollection<ShapeItem> _listOfSelectedShapes = new ObservableCollection<ShapeItem>();

        #endregion

        // This is the "real" list that we will add/remove items from
        private readonly ObservableCollection<ICQueueSummary > _ListOfIcEmailSummaries = new ObservableCollection<ICQueueSummary >();
        // This is a read-only wrapper around the real list that we will expose to the UI
        private readonly ReadOnlyObservableCollection<ICQueueSummary > _readOnly__ListOfIcEmailSummaries;

        private List<EmailItem> _emailsForCurrentQueue;  

        private ICQueueSummary  _selectedICEmailSummary;

        private IcGisTypeSetting IcGisTypeSetting { get; set; }


        #endregion

        #region Constructor
        protected Dockpane_IC_LoaderViewModel()
        {
           
            // Create the public, read-only collection that the UI will bind to
            _readOnly__ListOfIcEmailSummaries = new ReadOnlyObservableCollection<ICQueueSummary >(_ListOfIcEmailSummaries);

            // This is a key step from the sample. It allows a background thread to safely update a collection that the UI is bound to.
            BindingOperations.EnableCollectionSynchronization(_readOnly__ListOfIcEmailSummaries, _lockQueueCollection);

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

        // Public Read-Only wrappers for the UI to bind to safely
        public ReadOnlyObservableCollection<ICQueueSummary> ICEmailTypeSummaries { get; }
        public ReadOnlyObservableCollection<ShapeItem> ShapesToReview { get; }
        public ReadOnlyObservableCollection<ShapeItem> SelectedShapes { get; }

        // Properties for UI state
        private bool _isUIEnabled = false;
        public bool IsUIEnabled { get => _isUIEnabled; set => SetProperty(ref _isUIEnabled, value); }

        private string _statusMessage = "Please open a map view to use this tool.";
        public string StatusMessage { get => _statusMessage; set => SetProperty(ref _statusMessage, value); }

        // Properties for the currently selected items in the UI
        private ICQueueSummary _SelectedIcType;
        public ICQueueSummary SelectedIcType
        {
            get => _SelectedIcType;
            set
            {
                SetProperty(ref _SelectedIcType, value);
                if (_SelectedIcType != null && IsUIEnabled)
                {
                    _ = LoadEmailsForQueueAsync(_SelectedIcType);
                }
            }
        }

        // This property holds the SINGLE email we are currently processing
        private EmailItem _currentEmail;
        public EmailItem CurrentEmail
        {
            get => _currentEmail;
            set => SetProperty(ref _currentEmail, value);
        }

        #endregion



    }
}