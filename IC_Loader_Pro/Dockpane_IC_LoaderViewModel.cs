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



        //private ShapeItem _selectedShapeForReview;
        //public ShapeItem SelectedShapeForReview
        //{
        //    get => _selectedShapeForReview;
        //    set => SetProperty(ref _selectedShapeForReview, value);
        //}

        //private ShapeItem _selectedShapeToUse;
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
        public void SelectShapeFromTool(string elementName)
        {
            Log.RecordMessage($"SelectShapeFromTool received element name: '{elementName}'", BIS_Log.BisLogMessageType.Note);
            // Try to parse the Shape's ID from the element's Name property.
            if (int.TryParse(elementName, out int refId))
            {
                // Find the matching ShapeItem in our ViewModel's collections.
                var shapeToSelect = _shapesToReview.FirstOrDefault(s => s.ShapeReferenceId == refId) ??
                                    _selectedShapes.FirstOrDefault(s => s.ShapeReferenceId == refId);

                if (shapeToSelect != null)
                {
                    Log.RecordMessage($"Found matching ShapeItem with ID: {refId}. Updating UI.", BIS_Log.BisLogMessageType.Note);
                    // Update the UI selections on the main thread.
                    FrameworkApplication.Current.Dispatcher.Invoke(() =>
                    {
                        if (_shapesToReview.Contains(shapeToSelect))
                        {
                            SelectedShapesToUse.Clear();
                            if (!SelectedShapesForReview.Contains(shapeToSelect))
                            {
                                SelectedShapesForReview.Clear();
                                SelectedShapesForReview.Add(shapeToSelect);
                                SelectedShapeForReview = shapeToSelect;
                            }
                        }
                        else if (_selectedShapes.Contains(shapeToSelect))
                        {
                            SelectedShapesForReview.Clear();
                            if (!SelectedShapesToUse.Contains(shapeToSelect))
                            {
                                SelectedShapesToUse.Clear();
                                SelectedShapesToUse.Add(shapeToSelect);
                                SelectedShapeToUse = shapeToSelect;
                            }
                        }
                    });
                }
                else
                {
                    Log.RecordMessage($"No matching ShapeItem found for ID: {refId}.", BIS_Log.BisLogMessageType.Warning);
                }
            }
            else
            {
                Log.RecordMessage($"Failed to parse Shape ID from element name: '{elementName}'.", BIS_Log.BisLogMessageType.Warning);
            }
        }

        #endregion
    }
}