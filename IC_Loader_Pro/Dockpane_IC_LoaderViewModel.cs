using ArcGIS.Core.CIM;
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
using System.Collections.Generic;
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
        #endregion

        #region Constructor
        protected Dockpane_IC_LoaderViewModel()
        {
           
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


            // This is a key step. It allows a background thread to safely update a collection that the UI is bound to.
            BindingOperations.EnableCollectionSynchronization(_readOnlyListOfQueues, _lockQueueCollection);
            BindingOperations.EnableCollectionSynchronization(ShapesToReview, _lock);
            BindingOperations.EnableCollectionSynchronization(SelectedShapes, _lock);

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
            ReloadFileSetCommand = new RelayCommand(async (param) => await OnReloadFileSetAsync(param as FileSetViewModel),(param) => param is FileSetViewModel);
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
            if (_isRefreshingShapes) return;

            try
            {
                _isRefreshingShapes = true;

                // The UI thread must acquire the lock before modifying the collections
                lock (_lock)
                {
                    _shapesToReview.Clear();
                    _selectedShapes.Clear();

                    var fileSetLookup = _foundFileSets.ToDictionary(fs => fs.FileName);

                    foreach (var shape in _allProcessedShapes)
                    {
                        if (fileSetLookup.TryGetValue(shape.SourceFile, out var parentFileSet))
                        {
                            if (!parentFileSet.ShowInMap)
                            {
                                continue;
                            }

                            // --- THIS IS THE CORRECTED FILTER LOGIC ---
                            if (parentFileSet.UseFilter)
                            {
                                // If filter is ON, only show auto-selected shapes.
                                if (shape.IsAutoSelected)
                                {
                                    _selectedShapes.Add(shape);
                                }
                                // Any shape that is not auto-selected is now hidden.
                            }
                            else
                            {
                                // If filter is OFF, separate shapes normally.
                                if (shape.IsAutoSelected)
                                {
                                    _selectedShapes.Add(shape);
                                }
                                else
                                {
                                    _shapesToReview.Add(shape);
                                }
                            }
                        }
                    }
                } // The lock is released here

                await RedrawAllShapesOnMapAsync();
            }
            finally
            {
                _isRefreshingShapes = false;
            }
        }



        private async void FileSetViewModel_PropertyChanged(object sender, System.ComponentModel.PropertyChangedEventArgs e)
        {
            // If a Show or Filter checkbox changes, refresh everything.
            if (e.PropertyName == nameof(FileSetViewModel.ShowInMap) || e.PropertyName == nameof(FileSetViewModel.UseFilter))
            {
                // Use the busy flag to prevent the crash
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
                        else return;
                    }
                    else return;
                }

                var layerParams = new LayerCreationParams(new Uri(filePath));
                var fileUri = new Uri(filePath);

                // --- THIS IS THE CORRECTED LOGIC ---
                // Handle each file type appropriately
                switch (fs.filesetType.ToLowerInvariant())
                {
                    case "shapefile":
                        // The generic version for FeatureLayer can use LayerCreationParams
                        createdLayer = LayerFactory.Instance.CreateLayer<FeatureLayer>(layerParams, activeMap);
                        break;
                    case "dwg":
                        // The non-generic version for a GroupLayer needs the Uri directly
                        createdLayer = LayerFactory.Instance.CreateLayer(fileUri, activeMap);
                        break;
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
                });
            });
        }

        private void PerformCleanup()
        {
            // If there's a path from a previous run, delete that folder now.
            if (!string.IsNullOrEmpty(_pathForNextCleanup))
            {
                try
                {
                    if (Directory.Exists(_pathForNextCleanup))
                    {
                        Directory.Delete(_pathForNextCleanup, true);
                    }
                }
                catch (Exception ex)
                {
                    Log.RecordError($"Failed to delete previous temp folder: {_pathForNextCleanup}", ex, "PerformCleanup");
                }
                finally
                {
                    // Clear the path regardless of success or failure.
                    _pathForNextCleanup = null;
                }
            }
        }

        private async Task OnReloadFileSetAsync(FileSetViewModel fileSetVM)
        {
            Log.RecordMessage($"Reload requested for fileset: {fileSetVM.FileName}", BisLogMessageType.Note);
            // TODO: This is a complex operation that requires refactoring the feature
            // processing logic to run for a single fileset.
            ArcGIS.Desktop.Framework.Dialogs.MessageBox.Show("Reload functionality is not yet implemented.", "Coming Soon");
            await Task.CompletedTask;
        }

        #endregion
    }
}