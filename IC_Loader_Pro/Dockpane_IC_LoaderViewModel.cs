using ArcGIS.Core.CIM;
using ArcGIS.Core.Geometry;
using ArcGIS.Desktop.Catalog;
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
using IC_Loader_Pro.Services;
using IC_Loader_Pro.ViewModels;
using IC_Loader_Pro.Views;
using IC_Rules_2025;
using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Data;
using System.Windows.Input;
using static BIS_Log;
using static IC_Loader_Pro.Module1;
using Action = System.Action;
using Exception = System.Exception; // For  Log

namespace IC_Loader_Pro
{
    internal partial class  Dockpane_IC_LoaderViewModel : DockPane
    {
        #region Private Members
        /// <summary>
        /// The unique ID of this dockpane, must match the ID in Config.daml
        /// </summary>
        public const string DockPaneId = "IC_Loader_Pro_Dockpane_IC_Loader";
        public const string SelectToolId = "IC_Loader_Pro_SelectShapeTool";

        private readonly object _lockQueueCollection = new object();
        // This is the "real" list that we will add/remove items from
        private readonly ObservableCollection<ICQueueSummary> _ListOfIcEmailTypeSummaries = new ObservableCollection<ICQueueSummary>();
        // This is a read-only wrapper around the real list that we will expose to the UI
        private readonly ReadOnlyObservableCollection<ICQueueSummary> _readOnlyListOfQueues;

        // This collection will hold the filesets for the currently active email
        private readonly ObservableCollection<FileSetViewModel> _foundFileSets = new ObservableCollection<FileSetViewModel>();
        public ReadOnlyObservableCollection<FileSetViewModel> _readOnlyFoundFileSets { get; }

        private readonly ObservableCollection<ShapeItem> _shapesToReview = new ObservableCollection<ShapeItem>();
        private readonly ObservableCollection<ShapeItem> _selectedShapes = new ObservableCollection<ShapeItem>();
        public ReadOnlyObservableCollection<ShapeItem> ShapesToReview { get; }
        public ReadOnlyObservableCollection<ShapeItem> SelectedShapes { get; }

        private List<ShapeItem> _allProcessedShapes = new List<ShapeItem>();

        private readonly List<Layer> _manuallyLoadedLayers = new List<Layer>();

        private string _pathForNextCleanup;

        private bool _isRefreshingShapes = false;

        private MapPoint _currentSiteLocation;
        private GraphicsLayer _graphicsLayer = null;
        private ICQueueSummary _selectedQueue;
        private bool _isInitialized = false;
        private readonly object _lock = new object();
        private string _statusMessage = "Please open or create an ArcGIS Pro project.";
        public string StatusMessage { get => _statusMessage; set => SetProperty(ref _statusMessage, value); }

        public ShapeItem SelectedShapeForReview { get; set; }
        public ShapeItem SelectedShapeToUse { get; set; }

        private EmailItem _currentEmail;
        private EmailClassificationResult _currentClassification;

        private List<IcTestResult> _currentFilesetTestResults;

        private string _currentEmailId;
        public string CurrentEmailId
        {
            get => _currentEmailId;
            set => SetProperty(ref _currentEmailId, value);
        }

        /// <summary>
        /// Holds the settings for the currently selected IC Type.
        /// </summary>
        private IcGisTypeSetting _currentIcSetting;

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

        private bool _isSelectToolActive = false;
        public bool IsSelectToolActive
        {
            get => _isSelectToolActive;
            set
            {
                if (SetProperty(ref _isSelectToolActive, value))
                {
                    ToggleSelectTool();
                }
            }
        }

        private bool _showInMap = true;
        public bool ShowInMap
        {
            get => _showInMap;
            set => SetProperty(ref _showInMap, value);
        }

        private bool _useFilter = true;
        public bool UseFilter
        {
            get => _useFilter;
            set => SetProperty(ref _useFilter, value);
        }

        private int _totalFeatureCount;
        public int TotalFeatureCount
        {
            get => _totalFeatureCount;
            set => SetProperty(ref _totalFeatureCount, value);
        }

        private int _filteredCount;
        public int FilteredCount
        {
            get => _filteredCount;
            set => SetProperty(ref _filteredCount, value);
        }

        private int _validFeatureCount;
        public int ValidFeatureCount
        {
            get => _validFeatureCount;
            set => SetProperty(ref _validFeatureCount, value);
        }

        private int _invalidFeatureCount;
        public int InvalidFeatureCount
        {
            get => _invalidFeatureCount;
            set => SetProperty(ref _invalidFeatureCount, value);
        }

        private bool _isInTestMode;
        public bool IsInTestMode
        {
            get => _isInTestMode;
            set
            {
                if (SetProperty(ref _isInTestMode, value))
                {
                    // Update the global flag
                    Module1.IsInTestMode = value;
                    // Refresh the email queues to reflect the new mode
                    _ = RefreshICQueuesAsync();
                }
            }
        }

        private double _zoomToSiteDistance = 500; // Default distance in feet
        public double ZoomToSiteDistance
        {
            get => _zoomToSiteDistance;
            set => SetProperty(ref _zoomToSiteDistance, value);
        }

        #endregion

        #region Constructor
        protected Dockpane_IC_LoaderViewModel()
        {
            _isInTestMode = Module1.IsInTestMode;
            this.Caption = $"IC Loader (Build: {Module1.BuildDate.ToLocalTime():yyyy-MM-dd HH:mm})";


            // Create the public, read-only collection that the UI will bind to
            _readOnlyListOfQueues = new ReadOnlyObservableCollection<ICQueueSummary>(_ListOfIcEmailTypeSummaries);
            _readOnlyFoundFileSets = new ReadOnlyObservableCollection<FileSetViewModel>(_foundFileSets);
            ShapesToReview = new ReadOnlyObservableCollection<ShapeItem>(_shapesToReview);
            SelectedShapes = new ReadOnlyObservableCollection<ShapeItem>(_selectedShapes);
            SelectedShapesForReview.CollectionChanged += SelectedShapesForReview_CollectionChanged;
            SelectedShapesToUse.CollectionChanged += SelectedShapesToUse_CollectionChanged;
            SelectedShapesForReview.CollectionChanged += OnSelectionChanged;
            SelectedShapesToUse.CollectionChanged += OnSelectionChanged;
            _foundFileSets.CollectionChanged += FoundFileSets_CollectionChanged;
            _selectedShapes.CollectionChanged += SelectedShapes_CollectionChanged;


            // This is a key step. It allows a background thread to safely update a collection that the UI is bound to.
            BindingOperations.EnableCollectionSynchronization(_readOnlyListOfQueues, _lockQueueCollection);
            BindingOperations.EnableCollectionSynchronization(ShapesToReview, _lock);
            BindingOperations.EnableCollectionSynchronization(SelectedShapes, _lock);

            // Initialize commands
            RefreshQueuesCommand = new RelayCommand(async () => await RefreshICQueuesAsync(), () => IsUIEnabled);
            SaveCommand = new RelayCommand(async () => await OnSave(),() => IsEmailActionEnabled && _selectedShapes.Any());
            SkipCommand = new RelayCommand(async () => await OnSkip(), () => IsEmailActionEnabled);
            RejectCommand = new RelayCommand(async () => await OnReject(), () => IsEmailActionEnabled);
            ShowNotesCommand = new RelayCommand(async () => await OnShowNotes(), () => IsUIEnabled);
            SearchCommand = new RelayCommand(async () => await OnSearch(), () => IsUIEnabled);
            ToolsCommand = new RelayCommand(async () => await OnTools(), () => IsUIEnabled);
            OptionsCommand = new RelayCommand(async () => await OnOptions(), () => IsUIEnabled);
            ShowResultsCommand = new RelayCommand(OnShowResults, () => _currentEmailTestResult != null);
            AddSelectedShapeCommand = new RelayCommand(AddSelectedShape, () => SelectedShapesForReview.Any());// The button is enabled only if SelectedShapeForReview is not null
            RemoveSelectedShapeCommand = new RelayCommand(RemoveSelectedShape, () => SelectedShapesToUse.Any()); // The button is enabled only if SelectedShapeToUse is not null
            AddAllShapesCommand = new RelayCommand(OnAddAllShapes, () => _shapesToReview.Any());
            RemoveAllShapesCommand = new RelayCommand(OnRemoveAllShapes, () => _selectedShapes.Any());
            ZoomToAllCommand = new RelayCommand(async () => await OnZoomToAllAsync(),() => _shapesToReview.Any() || _selectedShapes.Any());
            ZoomToSelectedReviewShapeCommand = new RelayCommand(async () => await OnZoomToSelectedReviewShape(), () => SelectedShapesForReview.Any());
            ZoomToSelectedUseShapeCommand = new RelayCommand(async () => await OnZoomToSelectedUseShape(), () => SelectedShapesToUse.Any());
            ZoomToSiteCommand = new RelayCommand(async () => await OnZoomToSiteAsync(),() => _currentSiteLocation != null);
            ClearSelectionCommand = new RelayCommand(OnClearSelection,() => SelectedShapesForReview.Any() || SelectedShapesToUse.Any());
            ActivateSelectToolCommand = new RelayCommand(ActivateSelectTool);
            HideSelectionCommand = new RelayCommand(async () => await OnHideSelectionAsync(),() => SelectedShapesForReview.Any() || SelectedShapesToUse.Any());
            UnhideAllCommand = new RelayCommand(async () => await OnUnhideAllAsync());
            LoadFileSetCommand = new RelayCommand(async (param) => await OnLoadFileSetAsync(param as FileSetViewModel),(param) => param is FileSetViewModel fs && !fs.IsLoadedInMap);
            ReloadFileSetCommand = new RelayCommand(async (param) => await OnReloadFileSetAsync(param as FileSetViewModel),(param) => param is FileSetViewModel fs && fs.IsLoadedInMap);
            AddSubmissionCommand = new RelayCommand(async () => await OnAddSubmissionAsync(), () => IsUIEnabled);
            CreateNewIcDeliverableCommand = new RelayCommand(async () => await OnCreateNewIcDeliverableAsync(), () => IsUIEnabled);
            OpenConnectionTesterCommand = new RelayCommand(OnOpenConnectionTester, () => IsUIEnabled);
            OpenEmailInOutlookCommand = new RelayCommand(async () => await OnOpenEmailInOutlook(), () => _currentEmail != null);
            ProcessManualLayerCommand = new RelayCommand(async () => await OnProcessManualLayerAsync(), () => IsEmailActionEnabled);
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

                if (value != null)
                {
                    _currentIcSetting = IcRules.ReturnIcGisTypeSettings(value.Name);
                    // The underscore discards the returned Task, which is a standard
                    _ = AddRequiredLayersToMapAsync();
                    _ = ProcessSelectedQueueAsync();
                }
                else
                {
                    _currentIcSetting = null;
                }             
            }
        }

        private async Task RefreshShapeListsAndMap()
        {
            // This is the ONLY method that manages the flag.
            if (_isRefreshingShapes) return;

            try
            {
                _isRefreshingShapes = true; // Set the flag

                Log.RecordMessage("--- Refreshing Shape Lists and Map ---", BisLogMessageType.Note);
                lock (_lock)
                {
                    _shapesToReview.Clear();
                    _selectedShapes.Clear();
                    var fileSetLookup = _foundFileSets.ToDictionary(fs => fs.FileName);
                    foreach (var shape in _allProcessedShapes)
                    {
                        if (fileSetLookup.TryGetValue(shape.SourceFile, out var parentFileSet))
                        {
                            if (!parentFileSet.ShowInMap) continue;
                            if (parentFileSet.UseFilter)
                            {
                                if (shape.IsAutoSelected) _selectedShapes.Add(shape);
                            }
                            else
                            {
                                if (shape.IsAutoSelected) _selectedShapes.Add(shape);
                                else _shapesToReview.Add(shape);
                            }
                        }
                    }
                }
                Log.RecordMessage($"--- Refresh Complete. Review: {_shapesToReview.Count}, Selected: {_selectedShapes.Count} ---", BisLogMessageType.Note);

                await RedrawAllShapesOnMapAsync();
            }
            finally
            {
                _isRefreshingShapes = false; // ALWAYS clear the flag
            }
        }

        //private async Task RefreshShapeListsAndMap()
        //{
        //    if (_isRefreshingShapes) return;

        //    try
        //    {
        //        _isRefreshingShapes = true;

        //        // The UI thread must acquire the lock before modifying the collections
        //        lock (_lock)
        //        {
        //            _shapesToReview.Clear();
        //            _selectedShapes.Clear();

        //            var fileSetLookup = _foundFileSets.ToDictionary(fs => fs.FileName);

        //            foreach (var shape in _allProcessedShapes)
        //            {
        //                if (fileSetLookup.TryGetValue(shape.SourceFile, out var parentFileSet))
        //                {
        //                    if (!parentFileSet.ShowInMap)
        //                    {
        //                        continue;
        //                    }

        //                    // --- THIS IS THE CORRECTED FILTER LOGIC ---
        //                    if (parentFileSet.UseFilter)
        //                    {
        //                        // If filter is ON, only show auto-selected shapes.
        //                        if (shape.IsAutoSelected)
        //                        {
        //                            _selectedShapes.Add(shape);
        //                        }
        //                        // Any shape that is not auto-selected is now hidden.
        //                    }
        //                    else
        //                    {
        //                        // If filter is OFF, separate shapes normally.
        //                        if (shape.IsAutoSelected)
        //                        {
        //                            _selectedShapes.Add(shape);
        //                        }
        //                        else
        //                        {
        //                            _shapesToReview.Add(shape);
        //                        }
        //                    }
        //                }
        //            }
        //        } // The lock is released here

        //        await RedrawAllShapesOnMapAsync();
        //    }
        //    finally
        //    {
        //        _isRefreshingShapes = false;
        //    }
        //}

        private void UpdateFileSetCounts()
        {
            // Group all processed shapes by their source file
            var shapesByFile = _allProcessedShapes.GroupBy(s => s.SourceFile);

            // First, reset counts for all filesets in case one was emptied
            foreach (var fsVM in _foundFileSets)
            {
                fsVM.TotalFeatureCount = 0;
                fsVM.FilteredCount = 0;
                fsVM.ValidFeatureCount = 0;
                fsVM.InvalidFeatureCount = 0;
            }

            // Now, calculate and set the new counts
            foreach (var group in shapesByFile)
            {
                var fileSetVM = _foundFileSets.FirstOrDefault(fs => fs.FileName == group.Key);
                if (fileSetVM != null)
                {
                    fileSetVM.TotalFeatureCount = group.Count();
                    fileSetVM.FilteredCount = group.Count(s => s.IsAutoSelected);
                    fileSetVM.ValidFeatureCount = group.Count(s => s.IsValid);
                    fileSetVM.InvalidFeatureCount = group.Count(s => !s.IsValid);
                }
            }
        }



        private async void FileSetViewModel_PropertyChanged(object sender, System.ComponentModel.PropertyChangedEventArgs e)
        {
            // If a Show or Filter checkbox changes, try to refresh.
            if (e.PropertyName == nameof(FileSetViewModel.ShowInMap) || e.PropertyName == nameof(FileSetViewModel.UseFilter))
            {
                // This method now ONLY checks the flag. It does not set it.
                if (_isRefreshingShapes) return;
                await RefreshShapeListsAndMap();
            }
        }

        private void FoundFileSets_CollectionChanged(object sender, System.Collections.Specialized.NotifyCollectionChangedEventArgs e)
        {
            if (e.NewItems != null)
            {
                foreach (FileSetViewModel item in e.NewItems)
                {
                    item.PropertyChanged += FileSetViewModel_PropertyChanged;
                }
            }
            if (e.OldItems != null)
            {
                foreach (FileSetViewModel item in e.OldItems)
                {
                    item.PropertyChanged -= FileSetViewModel_PropertyChanged;
                }
            }
        }

        private void SelectedShapes_CollectionChanged(object sender, System.Collections.Specialized.NotifyCollectionChangedEventArgs e)
        {
            // Tell the Save button to re-evaluate its enabled/disabled state
            (SaveCommand as RelayCommand)?.RaiseCanExecuteChanged();
        }


        private void SelectedShapesForReview_CollectionChanged(object sender, System.Collections.Specialized.NotifyCollectionChangedEventArgs e)
        {
            // When the selection changes, tell the command to re-evaluate its state.
            (AddSelectedShapeCommand as RelayCommand)?.RaiseCanExecuteChanged();
        }

        private void SelectedShapesToUse_CollectionChanged(object sender, System.Collections.Specialized.NotifyCollectionChangedEventArgs e)
        {
            (RemoveSelectedShapeCommand as RelayCommand)?.RaiseCanExecuteChanged();
        }

        /// <summary>
        /// This method is called whenever an item is added or removed from either of the
        /// selection collections.
        /// </summary>
        private async void OnSelectionChanged(object sender, System.Collections.Specialized.NotifyCollectionChangedEventArgs e)
        {
            // Update the map highlight and also re-evaluate the button commands.
            await UpdateMapHighlightAsync();
            (AddSelectedShapeCommand as RelayCommand)?.RaiseCanExecuteChanged();
            (RemoveSelectedShapeCommand as RelayCommand)?.RaiseCanExecuteChanged();
            (ZoomToSelectedReviewShapeCommand as RelayCommand)?.RaiseCanExecuteChanged();
            (ZoomToSelectedUseShapeCommand as RelayCommand)?.RaiseCanExecuteChanged();
            (ClearSelectionCommand as RelayCommand)?.RaiseCanExecuteChanged();
        }



        private ShapeItem _selectedShapeForReview;
        //public ShapeItem SelectedShapeForReview
        //{
        //    get => _selectedShapeForReview;
        //    set => SetProperty(ref _selectedShapeForReview, value);
        //}

        private ShapeItem _selectedShapeToUse;
        //public ShapeItem SelectedShapeToUse
        //{
        //    get => _selectedShapeToUse;
        //    set => SetProperty(ref _selectedShapeToUse, value);
        //}

        /// <summary>
        /// A collection of the shapes currently selected in the "Shapes to Review" list.
        /// </summary>
        public ObservableCollection<object> SelectedShapesForReview { get; } = new ObservableCollection<object>();

        /// <summary>
        /// A collection of the shapes currently selected in the "Selected Shapes to Use" list.
        /// </summary>
        public ObservableCollection<object> SelectedShapesToUse { get; } = new ObservableCollection<object>();

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
                _ = LoadAndInitializeAsync();

                // Set the flag so we don't run this full initialization again
                // every time the user switches between maps.
                _isInitialized = true;
            }
        }     



        private async Task OnAddAllShapes()
        {
            await RunOnUIThread(() =>
            {
                // To avoid modifying the collection while looping, create a temporary copy.
                var itemsToMove = _shapesToReview.ToList();
                foreach (var item in itemsToMove)
                {
                    _selectedShapes.Add(item);
                    _shapesToReview.Remove(item);
                }
            });

            // After moving the items, redraw the map to update their symbols.
            await RedrawAllShapesOnMapAsync();
        }

        private async Task OnRemoveAllShapes()
        {
            await RunOnUIThread(() =>
            {
                var itemsToMove = _selectedShapes.ToList();
                foreach (var item in itemsToMove)
                {
                    _shapesToReview.Add(item);
                    _selectedShapes.Remove(item);
                }
            });

            // After moving the items, redraw the map to update their symbols.
            await RedrawAllShapesOnMapAsync();
        }

        private void OnClearSelection()
        {
            // This simply clears the collections that are bound to the
            // DataGrid's SelectedItems property, it doesn't move any items.
            RunOnUIThread(() =>
            {
                SelectedShapesForReview.Clear();
                SelectedShapesToUse.Clear();
            });
        }


        private async void ActivateSelectTool()
        {
            Log.RecordMessage("ActivateSelectTool command executed. Attempting to set current tool.", BIS_Log.BisLogMessageType.Note);
            await FrameworkApplication.SetCurrentToolAsync(SelectToolId);
        }



        /// <summary>
        /// Gathers all geometries from the review list, the selected list, and the site point,
        /// and then zooms the map to their combined extent.
        /// </summary>
        private async Task ZoomToAllAndSiteAsync()
        {
            // 1. Create a list to hold all the geometries we want to zoom to.
            var geometriesToZoom = new List<Geometry>();

            // 2. Add all the polygons from both the "Review" and "Use" lists.
            geometriesToZoom.AddRange(_shapesToReview.Select(s => s.Geometry));
            geometriesToZoom.AddRange(_selectedShapes.Select(s => s.Geometry));

            // 3. If a site location has been found, add it to the list as well.
            if (_currentSiteLocation != null)
            {
                geometriesToZoom.Add(_currentSiteLocation);
            }

            // 4. Call our existing generic zoom helper with the complete list.
            await ZoomToGeometryAsync(geometriesToZoom);
        }


        private async Task OnZoomToAllAsync()
        {
            await ZoomToAllAndSiteAsync();
        }

        private async Task ResetStateAsync()
        {
            // Clean up temporary files from the last run
            await PerformCleanupAsync();
            // Remove any layers that were manually loaded
            await ClearManuallyLoadedLayersAsync();

            // Reset all backing fields for the current deliverable
            _currentEmail = null;
            _currentClassification = null;
            _currentEmailTestResult = null;
            _currentAttachmentAnalysis = null;
            _currentSiteLocation = null;

            // Clear all the shape and fileset collections
            _allProcessedShapes.Clear();
            _foundFileSets.Clear();

            // Use the UI thread to clear collections bound to the UI
            await RunOnUIThread(() =>
            {
                _shapesToReview.Clear();
                _selectedShapes.Clear();
            });

            // Reset all the properties displayed in the UI
            CurrentEmailId = null;
            CurrentEmailSubject = "No email selected";
            CurrentPrefId = "N/A";
            CurrentAltId = "N/A";
            CurrentActivityNum = "N/A";
            CurrentDelId = "Pending";
            IsEmailActionEnabled = false;

            // Redraw the map to clear any old graphics
            await RedrawAllShapesOnMapAsync();
        }

        /// <summary>
        /// Zooms the active map view to the full extent of a collection of geometries.
        /// </summary>
        /// <param name="geometriesToZoomTo">A collection of Geometry objects to include in the extent.</param>
        private async Task ZoomToGeometryAsync(IEnumerable<Geometry> geometriesToZoomTo)
        {           
            // All map interactions must be run on the ArcGIS Pro main thread.
            // Filter out any null geometries
            var validGeometries = geometriesToZoomTo.Where(g => g != null && !g.IsEmpty).ToList();
            if (!validGeometries.Any()) return;

            await QueuedTask.Run(() =>
            {
                var mapView = MapView.Active;
                if (mapView == null) return;

                // --- THIS IS THE CORRECT METHOD ---
                // 1. Create a new EnvelopeBuilder, starting with the extent of the first geometry.
                var envelopeBuilder = new EnvelopeBuilderEx(validGeometries.First().Extent);

                // 2. Loop through the rest of the geometries and combine their extents.
                foreach (var geom in validGeometries.Skip(1))
                {
                    envelopeBuilder.Union(geom.Extent);
                }

                // 3. Get the final, combined envelope from the builder.
                var fullExtent = envelopeBuilder.ToGeometry();
                // ------------------------------------

                // Zoom to the combined extent with a small buffer (10% larger).
                mapView.ZoomTo(fullExtent.Expand(1.1, 1.1, true), TimeSpan.FromSeconds(0.5));
            });
        }

        /// <summary>
        /// Clears the highlight layer and draws the currently selected shapes with a highlight symbol.
        /// </summary>
        private async Task UpdateMapHighlightAsync()
        {
            await QueuedTask.Run(() =>
            {
                var mapView = MapView.Active;
                if (mapView == null) return;

                // Find our dedicated highlight layer
                var highlightLayer = mapView.Map.FindLayers("IC Loader Highlight").FirstOrDefault() as GraphicsLayer;
                if (highlightLayer == null) return;

                // Clear all previous highlights
                highlightLayer.RemoveElements();

                // 1. Create the stroke (the outline) for the polygon symbol.
                CIMStroke outline = SymbolFactory.Instance.ConstructStroke(
                    ColorFactory.Instance.CreateRGBColor(255, 255, 0), // Yellow  color
                    3.0, // 3-point width
                    SimpleLineStyle.Solid);

                // 2. Now, construct the polygon symbol using the fill color and the outline stroke.
                CIMPolygonSymbol highlightSymbol = SymbolFactory.Instance.ConstructPolygonSymbol(ColorFactory.Instance.CreateRGBColor(255, 255, 0, 30), SimpleFillStyle.Solid,outline);               

                // Get the geometries from BOTH selection lists
                var selectedGeometries = SelectedShapesForReview.OfType<ShapeItem>().Select(s => s.Geometry)
                                             .Concat(SelectedShapesToUse.OfType<ShapeItem>().Select(s => s.Geometry)).ToList();

                // Add each selected shape's geometry to the highlight layer
                foreach (var geom in selectedGeometries)
                {
                    if (geom != null)
                    {
                        highlightLayer.AddElement(geom, highlightSymbol);
                    }
                }
            });
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
            if (IsOnUIThread)
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
        private bool IsOnUIThread
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

        #region Public Methods for Tool Interaction

        /// <summary>
        /// A helper for the SelectShapeTool to get a reference to the graphics layer.
        /// </summary>
        public GraphicsLayer GetGraphicsLayer()
        {
            return _graphicsLayer;
        }

        /// <summary>
        /// Processes a shape selection coming from the custom map tool.
        /// </summary>
        public async void SelectShapeFromTool(string elementName)
        {
            if (int.TryParse(elementName, out int refId))
            {
                var shapeToSelect = _shapesToReview.FirstOrDefault(s => s.ShapeReferenceId == refId) ??
                                    _selectedShapes.FirstOrDefault(s => s.ShapeReferenceId == refId);

                if (shapeToSelect != null)
                {
                    // The logic is now very simple. We just update the collection.
                    // The new two-way behavior will see this change and force the UI to update visually.
                    await RunOnUIThread(() =>
                    {
                        if (_shapesToReview.Contains(shapeToSelect))
                        {
                            SelectedShapesToUse.Clear();
                            SelectedShapesForReview.Clear();
                            SelectedShapesForReview.Add(shapeToSelect);
                        }
                        else if (_selectedShapes.Contains(shapeToSelect))
                        {
                            SelectedShapesForReview.Clear();
                            SelectedShapesToUse.Clear();
                            SelectedShapesToUse.Add(shapeToSelect);
                        }
                    });
                }
            }
        }

        private async void ToggleSelectTool()
        {
            if (IsSelectToolActive)
            {
                await FrameworkApplication.SetCurrentToolAsync(SelectToolId);
            }
            else
            {
                await FrameworkApplication.SetCurrentToolAsync("esri_mapping_exploreTool");
            }
        }

        public void DeactivateSelectTool()
        {
            // This is called by the tool after a successful click
            IsSelectToolActive = false;
        }

        public void ToggleShapeSelectionFromTool(string elementName)
        {
            if (int.TryParse(elementName, out int refId))
            {
                var shapeToToggle = _shapesToReview.FirstOrDefault(s => s.ShapeReferenceId == refId) ??
                                    _selectedShapes.FirstOrDefault(s => s.ShapeReferenceId == refId);

                if (shapeToToggle != null)
                {
                    FrameworkApplication.Current.Dispatcher.Invoke(() =>
                    {
                        if (_shapesToReview.Contains(shapeToToggle))
                        {
                            // Logic for the "ToReview" list
                            if (SelectedShapesForReview.Contains(shapeToToggle))
                            {
                                SelectedShapesForReview.Remove(shapeToToggle);
                            }
                            else
                            {
                                SelectedShapesForReview.Add(shapeToToggle);
                            }
                        }
                        else if (_selectedShapes.Contains(shapeToToggle))
                        {
                            // Logic for the "ToUse" list
                            if (SelectedShapesToUse.Contains(shapeToToggle))
                            {
                                SelectedShapesToUse.Remove(shapeToToggle);
                            }
                            else
                            {
                                SelectedShapesToUse.Add(shapeToToggle);
                            }
                        }
                    });
                }
            }
        }

        /// <summary>
        /// Sets the IsHidden flag on all currently selected shapes and redraws the map.
        /// </summary>
        private async Task OnHideSelectionAsync()
        {
            // Combine the selected items from both lists
            var itemsToHide = SelectedShapesForReview.OfType<ShapeItem>()
                                .Concat(SelectedShapesToUse.OfType<ShapeItem>()).ToList();

            foreach (var item in itemsToHide)
            {
                item.IsHidden = true;
            }

            // Clear the selection from the UI
            SelectedShapesForReview.Clear();
            SelectedShapesToUse.Clear();

            // Redraw the map to reflect the changes
            await RedrawAllShapesOnMapAsync();
        }

        /// <summary>
        /// Clears the IsHidden flag on ALL shapes and redraws the map.
        /// </summary>
        private async Task OnUnhideAllAsync()
        {
            // Combine all items from both lists
            var allItems = _shapesToReview.Concat(_selectedShapes);
            foreach (var item in allItems)
            {
                item.IsHidden = false;
            }
            await RedrawAllShapesOnMapAsync();
        }

        private async Task OnLoadFileSetAsync(FileSetViewModel fileSetVM)
        {
            if (fileSetVM == null) return;

            // This will hold the layer we create, regardless of its specific type (FeatureLayer or GroupLayer)
            Layer createdLayer = null;

            // Part 1: Perform GIS work on the background thread (QueuedTask)
            await QueuedTask.Run(() =>
            {
                var activeMap = MapView.Active?.Map;
                if (activeMap == null) return;

                var fs = fileSetVM.Model;
                string extension;
                switch (fs.filesetType.ToLowerInvariant())
                {
                    case "shapefile": extension = "shp"; break;
                    case "dwg": extension = "dwg"; break;
                    default: extension = fs.filesetType; break;
                }
                string filePath = Path.Combine(fs.path, fs.fileName + "." + extension);

                if (!File.Exists(filePath))
                {
                    // Fallback search logic
                    if (_pathForNextCleanup != null)
                    {
                        var files = Directory.GetFiles(_pathForNextCleanup, fs.fileName + "." + extension, SearchOption.AllDirectories);
                        if (files.Any()) filePath = files.First();
                        else
                        {
                            Log.RecordError($"Could not find the file to load: {filePath}", null, "OnLoadFileSetAsync");
                            return;
                        }
                    }
                    else return;
                }

                try // This try block will catch the crash.
                {
                    var layerParams = new LayerCreationParams(new Uri(filePath));
                    var fileUri = new Uri(filePath);

                    switch (fs.filesetType.ToLowerInvariant())
                    {
                        case "shapefile":
                            createdLayer = LayerFactory.Instance.CreateLayer<FeatureLayer>(layerParams, activeMap);
                            break;
                        case "dwg":
                            createdLayer = LayerFactory.Instance.CreateLayer(fileUri, activeMap);
                            break;
                    }
                }
                catch (Exception ex)
                {
                    // If an error occurs, log it and inform the user gracefully.
                    Log.RecordError($"Failed to load layer from source: {filePath}", ex, "OnLoadFileSetAsync");
                    ArcGIS.Desktop.Framework.Dialogs.MessageBox.Show(
                        $"Could not load the selected layer. The data source may be invalid or incomplete.\n\nError: {ex.Message}",
                        "Layer Load Error");
                    createdLayer = null; // Ensure createdLayer is null on failure.
                }
            });

            // Part 2: Perform UI updates back on the main UI thread
            if (createdLayer != null)
            {
                _manuallyLoadedLayers.Add(createdLayer);

                await FrameworkApplication.Current.Dispatcher.InvokeAsync(() =>
                {
                    fileSetVM.IsLoadedInMap = true;
                    (LoadFileSetCommand as RelayCommand)?.RaiseCanExecuteChanged();
                    (ReloadFileSetCommand as RelayCommand)?.RaiseCanExecuteChanged();
                });
            }
        }

        private Task ClearManuallyLoadedLayersAsync()
        {
            // Part 1: Perform GIS work on the background thread
            return QueuedTask.Run(() =>
            {
                var activeMap = MapView.Active?.Map;
                if (activeMap == null || !_manuallyLoadedLayers.Any()) return;

                activeMap.RemoveLayers(_manuallyLoadedLayers);
                _manuallyLoadedLayers.Clear();

                // Part 2: Dispatch UI updates back to the main UI thread
                FrameworkApplication.Current.Dispatcher.Invoke(() =>
                {
                    foreach (var fsVM in _foundFileSets)
                    {
                        fsVM.IsLoadedInMap = false;
                    }
                    (LoadFileSetCommand as RelayCommand)?.RaiseCanExecuteChanged();
                    (ReloadFileSetCommand as RelayCommand)?.RaiseCanExecuteChanged();
                });
            });
        }

        private async Task PerformCleanupAsync()
        {
            if (string.IsNullOrEmpty(_pathForNextCleanup)) return;

            string pathToClean = _pathForNextCleanup;
            _pathForNextCleanup = null; // Clear the field immediately

            if (!Directory.Exists(pathToClean)) return;

            // --- START OF MODIFIED LOGIC ---
            // Run this cleanup on a background thread to avoid blocking the UI
            await Task.Run(() =>
            {
                for (int i = 0; i < 5; i++) // Retry up to 5 times
                {
                    try
                    {
                        // 1. Get all files in the directory and all its subdirectories.
                        var files = Directory.GetFiles(pathToClean, "*", SearchOption.AllDirectories);

                        // 2. Try to delete each file individually.
                        foreach (var file in files)
                        {
                            try
                            {
                                File.SetAttributes(file, FileAttributes.Normal); // Ensure file is not read-only
                                File.Delete(file);
                            }
                            catch (IOException)
                            {
                                // This file is likely locked by ArcGIS Pro. Silently ignore it.
                            }
                        }

                        // 3. After attempting to delete all files, delete the main directory.
                        //    The 'true' parameter means it will also delete any (now empty) subdirectories.
                        Directory.Delete(pathToClean, true);

                        Log.RecordMessage($"Successfully cleaned up temp folder: {pathToClean}", BisLogMessageType.Note);
                        return; // Exit successfully if the directory is deleted
                    }
                    catch (IOException)
                    {
                        // This will be caught if the directory still contains a locked file.
                        // We will wait and retry.
                        Task.Delay(300).Wait(); // Wait a bit longer before retrying
                    }
                    catch (Exception ex)
                    {
                        // Catch any other unexpected errors, log them once, and stop trying.
                        Log.RecordError($"An unexpected error occurred during cleanup of {pathToClean}.", ex, "PerformCleanupAsync");
                        return;
                    }
                }

                // If we exit the loop, it means the folder still couldn't be deleted.
                // Log this just once as a warning instead of a verbose error.
                Log.RecordMessage($"Could not fully clean up temp folder {pathToClean} due to persistent file locks.", BisLogMessageType.Warning);
            });
            // --- END OF MODIFIED LOGIC ---
        }

        private async Task OnReloadFileSetAsync(FileSetViewModel fileSetVM)
        {
            if (fileSetVM == null || _isRefreshingShapes) return;

            Log.RecordMessage($"Reload requested for fileset: {fileSetVM.FileName}", BisLogMessageType.Note);

            // This method no longer needs its own try/finally or to manage the busy flag.

            // 1. If the layer is on the map, remove it to invalidate Pro's cache.
            if (fileSetVM.IsLoadedInMap)
            {
                await QueuedTask.Run(() =>
                {
                    var layerToRemove = _manuallyLoadedLayers.FirstOrDefault(l => l.Name.Equals(fileSetVM.Model.fileName, StringComparison.OrdinalIgnoreCase));
                    if (layerToRemove != null)
                    {
                        MapView.Active?.Map.RemoveLayer(layerToRemove);
                        _manuallyLoadedLayers.Remove(layerToRemove);
                    }
                });

                await FrameworkApplication.Current.Dispatcher.InvokeAsync(() =>
                {
                    fileSetVM.IsLoadedInMap = false;
                    (LoadFileSetCommand as RelayCommand)?.RaiseCanExecuteChanged();
                    (ReloadFileSetCommand as RelayCommand)?.RaiseCanExecuteChanged();
                });
            }

            // 2. Remove all old shapes for this fileset from the master list.
            _allProcessedShapes.RemoveAll(s => s.SourceFile == fileSetVM.FileName);

            // 3. Re-process the single fileset.
            var namedTests = new IcNamedTests(Module1.Log, Module1.PostGreTool);
            var featureService = new FeatureProcessingService(Module1.IcRules, namedTests, Module1.Log);
            var reloadTestResult = namedTests.returnNewTestResult("GIS_Root_Email_Load", fileSetVM.FileName, IcTestResult.TestType.Submission);

            List<ShapeItem> reloadedShapes = await featureService.AnalyzeFeaturesFromFilesetsAsync(
                new List<fileset> { fileSetVM.Model },
                SelectedIcType.Name,
                _currentSiteLocation,
                reloadTestResult);

            // 4. Add the newly processed shapes back to the master list.
            if (reloadedShapes.Any())
            {
                _allProcessedShapes.AddRange(reloadedShapes);
            }

            // 5. Update the counts in the UI.
            UpdateFileSetCounts();

            // --- THIS IS THE CORRECTED LINE ---
            // Use BeginInvoke to schedule the refresh and break the async context,
            // which can solve stubborn UI update issues.
            //FrameworkApplication.Current.Dispatcher.BeginInvoke( new Action(async () => await RefreshShapeListsAndMap()));
            await RefreshShapeListsAndMap();
        }

        //private async Task OnReloadFileSetAsync(FileSetViewModel fileSetVM)
        //{
        //    // We only check the flag here to prevent starting a reload if another refresh is already running.
        //    if (fileSetVM == null || _isRefreshingShapes) return;

        //    Log.RecordMessage($"Reload requested for fileset: {fileSetVM.FileName}", BisLogMessageType.Note);

        //    // 1. If the layer is on the map, remove it to invalidate Pro's cache.
        //    if (fileSetVM.IsLoadedInMap)
        //    {
        //        await QueuedTask.Run(() =>
        //        {
        //            var layerToRemove = _manuallyLoadedLayers.FirstOrDefault(l => l.Name.Equals(fileSetVM.Model.fileName, StringComparison.OrdinalIgnoreCase));
        //            if (layerToRemove != null)
        //            {
        //                MapView.Active?.Map.RemoveLayer(layerToRemove);
        //                _manuallyLoadedLayers.Remove(layerToRemove);
        //            }
        //        });

        //        await FrameworkApplication.Current.Dispatcher.InvokeAsync(() =>
        //        {
        //            fileSetVM.IsLoadedInMap = false;
        //            (LoadFileSetCommand as RelayCommand)?.RaiseCanExecuteChanged();
        //            (ReloadFileSetCommand as RelayCommand)?.RaiseCanExecuteChanged();
        //        });
        //    }

        //    // 2. Remove all old shapes for this fileset from the master list.
        //    _allProcessedShapes.RemoveAll(s => s.SourceFile == fileSetVM.FileName);

        //    // 3. Re-process the single fileset.
        //    var namedTests = new IcNamedTests(Module1.Log, Module1.PostGreTool);
        //    var featureService = new FeatureProcessingService(Module1.IcRules, namedTests, Module1.Log);
        //    var reloadTestResult = namedTests.returnNewTestResult("GIS_Root_Email_Load", fileSetVM.FileName, IcTestResult.TestType.Submission);

        //    List<ShapeItem> reloadedShapes = await featureService.AnalyzeFeaturesFromFilesetsAsync(
        //        new List<fileset> { fileSetVM.Model },
        //        SelectedIcType.Name,
        //        _currentSiteLocation,
        //        reloadTestResult);

        //    // 4. Add the newly processed shapes back to the master list.
        //    if (reloadedShapes.Any())
        //    {
        //        _allProcessedShapes.AddRange(reloadedShapes);
        //    }

        //    // 5. Update the counts in the UI.
        //    UpdateFileSetCounts();

        //    // 6. Refresh the data grids and the map. 

        //    await RefreshShapeListsAndMap();
        //}

        private async Task OnAddSubmissionAsync()
        {
            //var browseFilter = new BrowseProjectFilter();//("esri_browseDialogFilters_shapefiles_all", "esri_browseDialogFilters_cad_all")
            //browseFilter.AddCanBeTypeId("shapefile_general");
            //browseFilter.AddCanBeTypeId("cad_general");
            //{
            //    Name = "GIS Files (Shapefile, DWG)" // This name appears in the file type dropdown
            //};
            // Use ArcGIS Pro's Open Item dialog for a native look and feel
            var openDialog = new OpenItemDialog
            {
                Title = "Add Submission Fileset",
                MultiSelect = false,
                //BrowseFilter = browseFilter
            };

            if (openDialog.ShowDialog() != true)
            {
                return; // User canceled
            }

            // Get the selected item (which could be a shapefile or a DWG)
            var selectedItem = openDialog.Items.FirstOrDefault();
            if (selectedItem == null) return;

            Log.RecordMessage($"User selected file to add: {selectedItem.Path}", BisLogMessageType.Note);

            // We need to copy the entire fileset to our current email's temp folder
            // to ensure it's processed and cleaned up correctly.
            if (string.IsNullOrEmpty(_pathForNextCleanup))
            {
                // If no email is active, we can't add a submission.
                ArcGIS.Desktop.Framework.Dialogs.MessageBox.Show("Please process an email before adding a manual submission.", "No Active Email");
                return;
            }

            try
            {
               // _isRefreshingShapes = true;

                // Create a fileset object representing the source data
                var sourceFileSet = Module1.IcRules.ReturnFileSetsFromDirectory_NewMethod(Path.GetDirectoryName(selectedItem.Path),"",false)
                                        .FirstOrDefault(fs => fs.fileName.Equals(Path.GetFileNameWithoutExtension(selectedItem.Name), StringComparison.OrdinalIgnoreCase));

                if (sourceFileSet == null)
                {
                    ArcGIS.Desktop.Framework.Dialogs.MessageBox.Show($"Could not identify a valid fileset for '{selectedItem.Name}'.", "Fileset Error");
                    return;
                }

                // Copy all component files of the fileset to our active temp directory
                foreach (var ext in sourceFileSet.extensions)
                {
                    string sourceFile = Path.Combine(sourceFileSet.path, $"{sourceFileSet.fileName}.{ext}");
                    string destFile = Path.Combine(_pathForNextCleanup, $"{sourceFileSet.fileName}.{ext}");
                    if (File.Exists(sourceFile))
                    {
                        File.Copy(sourceFile, destFile, true);
                    }
                }

                // Now, create a new fileset model pointing to the copied location
                var newFileSetInTemp = new fileset
                {
                    fileName = sourceFileSet.fileName,
                    filesetType = sourceFileSet.filesetType,
                    path = _pathForNextCleanup, // IMPORTANT: Use the active temp path
                    extensions = sourceFileSet.extensions,
                    validSet = sourceFileSet.validSet
                };

                // Create a view model for the new fileset and add it to the UI
                var fileSetVM = new FileSetViewModel(newFileSetInTemp)
                {
                    UseFilter = !newFileSetInTemp.filesetType.Equals("shapefile", StringComparison.OrdinalIgnoreCase)
                };
                _foundFileSets.Add(fileSetVM);

                // Re-process the newly added fileset
                var namedTests = new IcNamedTests(Module1.Log, Module1.PostGreTool);
                var featureService = new FeatureProcessingService(Module1.IcRules, namedTests, Module1.Log);
                var processTestResult = namedTests.returnNewTestResult("GIS_Root_Email_Load", fileSetVM.FileName, IcTestResult.TestType.Submission);

                List<ShapeItem> newShapes = await featureService.AnalyzeFeaturesFromFilesetsAsync(
                    new List<fileset> { newFileSetInTemp },
                    SelectedIcType.Name,
                    _currentSiteLocation,
                    processTestResult);

                // Add the new shapes to our master list and refresh everything
                if (newShapes.Any())
                {
                    _allProcessedShapes.AddRange(newShapes);
                }
                UpdateFileSetCounts();
                await RefreshShapeListsAndMap();
            }
            catch (Exception ex)
            {
                Log.RecordError("Failed to add manual submission.", ex, nameof(OnAddSubmissionAsync));
                ArcGIS.Desktop.Framework.Dialogs.MessageBox.Show($"An error occurred while adding the submission: {ex.Message}", "Error");
            }
            finally
            {
                _isRefreshingShapes = false;
            }
        }

        private async Task OnCreateNewIcDeliverableAsync()
        {
            // 1. Reset all state from any previous deliverable.
            await ResetStateAsync();

            // 1. Create an instance of the new window and its ViewModel
            var viewModel = new ViewModels.CreateIcDeliverableViewModel();
            var window = new CreateIcDeliverableWindow
            {
                DataContext = viewModel,
                Owner = FrameworkApplication.Current.MainWindow // This makes it a dialog of the main Pro window
            };

            // 2. Show the window and wait for the user to click "Create" or "Cancel"
            if (window.ShowDialog() != true)
            {
                return; // User canceled
            }

            // 3. If the user clicked "Create", get the data from the ViewModel
            string selectedIcType = viewModel.SelectedIcType;
            string prefId = viewModel.PrefId;
            string selectedFilePath = viewModel.GisFilePath;
            CurrentPrefId = prefId;

            Log.RecordMessage($"User initiated new deliverable. Type: {selectedIcType}, PrefID: {prefId}, File: {selectedFilePath}", BisLogMessageType.Note);

            // An active email process must exist to have a temp folder to copy into.
            // If not, we create a new one for this manual submission.
            if (string.IsNullOrEmpty(_pathForNextCleanup))
            {
                string localAppData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
                string addinTempRoot = Path.Combine(localAppData, "IC_Loader_Pro_Temp");
                Directory.CreateDirectory(addinTempRoot);
                _pathForNextCleanup = Path.Combine(addinTempRoot, Guid.NewGuid().ToString());
                Directory.CreateDirectory(_pathForNextCleanup);
            }

            try
            {
               // _isRefreshingShapes = true;

                var sourceFileSet = Module1.IcRules.ReturnFileSetsFromDirectory_NewMethod(Path.GetDirectoryName(selectedFilePath), "", false)
                                        .FirstOrDefault(fs => fs.fileName.Equals(Path.GetFileNameWithoutExtension(selectedFilePath), StringComparison.OrdinalIgnoreCase));

                if (sourceFileSet == null)
                {
                    ArcGIS.Desktop.Framework.Dialogs.MessageBox.Show($"Could not identify a valid fileset for '{Path.GetFileName(selectedFilePath)}'.", "Fileset Error");
                    return;
                }

                // ... The rest of the processing logic is the same as before ...
                foreach (var ext in sourceFileSet.extensions)
                {
                    string sourceFile = Path.Combine(sourceFileSet.path, $"{sourceFileSet.fileName}.{ext}");
                    string destFile = Path.Combine(_pathForNextCleanup, $"{sourceFileSet.fileName}.{ext}");
                    if (File.Exists(sourceFile)) File.Copy(sourceFile, destFile, true);
                }

                var newFileSetInTemp = new fileset
                {
                    fileName = sourceFileSet.fileName,
                    filesetType = sourceFileSet.filesetType,
                    path = _pathForNextCleanup,
                    extensions = sourceFileSet.extensions,
                    validSet = sourceFileSet.validSet
                };

                var fileSetVM = new FileSetViewModel(newFileSetInTemp)
                {
                    UseFilter = !newFileSetInTemp.filesetType.Equals("shapefile", StringComparison.OrdinalIgnoreCase)
                };
                _foundFileSets.Add(fileSetVM);

                var namedTests = new IcNamedTests(Module1.Log, Module1.PostGreTool);
                var featureService = new FeatureProcessingService(Module1.IcRules, namedTests, Module1.Log);
                // We need a site location for validation. Let's get it now.
                var siteLocation = await GetSiteCoordinatesFromNjemsAsync(prefId);
                var processTestResult = namedTests.returnNewTestResult("GIS_Root_Email_Load", fileSetVM.FileName, IcTestResult.TestType.Submission);

                List<ShapeItem> newShapes = await featureService.AnalyzeFeaturesFromFilesetsAsync(
                    new List<fileset> { newFileSetInTemp }, selectedIcType, siteLocation, processTestResult);

                if (newShapes.Any())
                {
                    _allProcessedShapes.AddRange(newShapes);
                }
                UpdateFileSetCounts();
                await RefreshShapeListsAndMap();

                IsEmailActionEnabled = true;
                (SaveCommand as RelayCommand)?.RaiseCanExecuteChanged();
            }
            catch (Exception ex)
            {
                Log.RecordError("Failed to add manual submission.", ex, nameof(OnCreateNewIcDeliverableAsync));
                ArcGIS.Desktop.Framework.Dialogs.MessageBox.Show($"An error occurred while adding the submission: {ex.Message}", "Error");
            }
            finally
            {
                _isRefreshingShapes = false;
            }
        }

        private void OnOpenConnectionTester()
        {
            var testerWindow = new Views.ConnectionTesterWindow
            {
                DataContext = new ConnectionTesterViewModel(),
                Owner = FrameworkApplication.Current.MainWindow
            };
            testerWindow.ShowDialog();
        }

        private async Task AddRequiredLayersToMapAsync()
        {
            if (_currentIcSetting == null || MapView.Active == null) return;

            // --- Part 1: Add the Proposed Feature Class ---
            var proposedFcRule = _currentIcSetting.ProposedFeatureClass;
            if (proposedFcRule != null && !string.IsNullOrEmpty(proposedFcRule.PostGreFeatureClassName))
            {
                await QueuedTask.Run(async () =>
                {
                    var activeMap = MapView.Active.Map;

                    // --- THIS IS THE CORRECTED CHECK (using EndsWith) ---
                    if (activeMap.GetLayersAsFlattenedList().OfType<FeatureLayer>()
                        .Any(fl => fl.GetFeatureClass()?.GetName().EndsWith(proposedFcRule.PostGreFeatureClassName, StringComparison.OrdinalIgnoreCase) == true))
                    {
                        return; // Layer from this source is already on the map
                    }

                    var gdbService = new Services.GeodatabaseService();
                    using (var proposedFc = await gdbService.GetFeatureClassAsync(proposedFcRule))
                    {
                        if (proposedFc != null)
                        {
                            var layerParams = new FeatureLayerCreationParams(proposedFc) { Name = proposedFcRule.PostGreFeatureClassName };
                            LayerFactory.Instance.CreateLayer<FeatureLayer>(layerParams, activeMap);
                        }
                    }
                });
            }

            // --- Part 2: Add the Shape Info Table ---
            var shapeInfoTableRule = _currentIcSetting.ShapeInfoTable;
            if (shapeInfoTableRule != null && !string.IsNullOrEmpty(shapeInfoTableRule.PostGreFeatureClassName))
            {
                await QueuedTask.Run(async () =>
                {
                    var activeMap = MapView.Active.Map;

                    // --- THIS IS THE CORRECTED CHECK (using EndsWith) ---
                    if (activeMap.GetStandaloneTablesAsFlattenedList()
                        .Any(t => t.GetTable()?.GetName().EndsWith(shapeInfoTableRule.PostGreFeatureClassName, StringComparison.OrdinalIgnoreCase) == true))
                    {
                        return; // Table from this source is already in the map
                    }

                    var gdbService = new Services.GeodatabaseService();
                    using (var shapeInfoTable = await gdbService.GetTableAsync(shapeInfoTableRule))
                    {
                        if (shapeInfoTable != null)
                        {
                            var tableParams = new StandaloneTableCreationParams(shapeInfoTable) { Name = shapeInfoTableRule.PostGreFeatureClassName };
                            StandaloneTableFactory.Instance.CreateStandaloneTable(tableParams, activeMap);
                        }
                    }
                });
            }
        }       
        #endregion
    }
}