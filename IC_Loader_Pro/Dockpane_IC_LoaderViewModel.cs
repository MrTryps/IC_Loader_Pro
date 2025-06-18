using ArcGIS.Core.Geometry;
using ArcGIS.Desktop.Core;
using ArcGIS.Desktop.Core.Events;
using ArcGIS.Desktop.Core.Geoprocessing;
using ArcGIS.Desktop.Framework;
using ArcGIS.Desktop.Framework.Contracts;
using ArcGIS.Desktop.Framework.Threading.Tasks;
using ArcGIS.Desktop.Mapping;
using IC_Loader_Pro.Models; // Your ICQueueInfo class
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
        private readonly ObservableCollection<ICQueueInfo> _listOfQueues = new ObservableCollection<ICQueueInfo>();

        // This is a read-only wrapper around the real list that we will expose to the UI
        private readonly ReadOnlyObservableCollection<ICQueueInfo> _readOnlyListOfQueues;

        private ICQueueInfo _selectedQueue;

        private bool _isInitialized = false;
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

        public bool IsUIEnabled { get; private set; }
        public string StatusMessage { get; private set; }


        #endregion

        #region ensure the project is ready
        public async Task LoadAndInitializeAsync()
        {
            // Ensure this complex initialization only ever runs once.

            {
                if (_isInitialized)
                    return;
                _isInitialized = true;
            }
            IsUIEnabled = true;

            SaveCommand = new RelayCommand(async () => await OnSave(), () => IsUIEnabled);
            SkipCommand = new RelayCommand(async () => await OnSkip(), () => IsUIEnabled);
            RejectCommand = new RelayCommand(async () => await OnReject(), () => IsUIEnabled);
            ShowNotesCommand = new RelayCommand(async () => await OnShowNotes(), () => IsUIEnabled);
            SearchCommand = new RelayCommand(async () => await OnSearch(), () => IsUIEnabled);
            ToolsCommand = new RelayCommand(async () => await OnTools(), () => IsUIEnabled);
            OptionsCommand = new RelayCommand(async () => await OnOptions(), () => IsUIEnabled);

            Log.recordMessage("Initializing Dockpane...", Bis_Log_Message_Type.Note);
            StatusMessage = "Initializing map and layers...";

            try
            {
                // We can now safely assume a map exists and is active.
                Map activeMap = MapView.Active?.Map;
                if (activeMap == null)
                {
                    Log.recordError("No active map found. Initialization cannot proceed.", null, nameof(LoadAndInitializeAsync));
                    StatusMessage = "Error: No active map found.";
                    return;
                }

                // Enforce the coordinate system on the active map.
                await QueuedTask.Run(() =>
                {
                    int requiredWkid = 2260;
                    if (activeMap.SpatialReference?.Wkid != requiredWkid)
                    {
                        var njStatePlane = SpatialReferenceBuilder.CreateSpatialReference(requiredWkid);
                        activeMap.SetSpatialReference(njStatePlane);
                    }
                });

                // Ensure our special "manually_added" scratch layer is ready.
                await EnsureManualAddLayerExistsAsync(activeMap);

                Log.recordMessage("Initialization complete.", Bis_Log_Message_Type.Note);
            }
            catch (Exception ex)
            {
                Log.recordError("A fatal error occurred during initialization.", ex, nameof(LoadAndInitializeAsync));
            }
        }


        private Task EnsureManualAddLayerExistsAsync(Map map)
        {
            return QueuedTask.Run(() =>
            {
                string layerName = "manually_added";
                int requiredWkid = 2260; // NAD 1983 NJ State Plane Feet

                var existingLayer = map.GetLayersAsFlattenedList()
                                        .FirstOrDefault(l => l.Name.Equals(layerName, StringComparison.CurrentCultureIgnoreCase)) as FeatureLayer;

                if (existingLayer != null)
                {
                    // The layer exists by name, now let's validate it thoroughly.
                    bool isLayerValid = false;
                    string validationError = "Unknown validation error.";

                    try
                    {
                        using (var featureClass = existingLayer.GetFeatureClass())
                        {
                            if (featureClass == null)
                            {
                                validationError = "Data source is broken or inaccessible.";
                            }
                            else
                            {
                                var definition = featureClass.GetDefinition();
                                if (definition.GetShapeType() != GeometryType.Polygon)
                                {
                                    validationError = "Geometry type is not Polygon.";
                                }
                                else if (definition.GetFields().Any(f => f.Name.Equals("id", StringComparison.CurrentCultureIgnoreCase)) == false)
                                {
                                    validationError = "Required 'id' field is missing.";
                                }
                                else if (definition.GetSpatialReference()?.Wkid != requiredWkid)
                                {
                                    validationError = $"Coordinate system is not the required NJ State Plane (WKID {requiredWkid}).";
                                }
                                else
                                {
                                    // If all checks pass, the layer is valid.
                                    isLayerValid = true;
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        validationError = $"An exception occurred while validating the layer: {ex.Message}";
                    }

                    if (isLayerValid)
                    {
                        Log.recordMessage($"Layer '{layerName}' already exists in the map and is valid.", Bis_Log_Message_Type.Note);
                        _manualAddLayer = existingLayer;
                        return; // We are done, the layer is good to use.
                    }
                    else
                    {
                        Log.recordMessage($"Removing invalid '{layerName}' layer. Reason: {validationError}", Bis_Log_Message_Type.Warning);
                        map.RemoveLayer(existingLayer);
                    }
                }

                // If we get to this point, it means the layer did not exist or was invalid and has been removed.
                string shapefilePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "IC_Loader_Pro", "manually_added.shp");
                Directory.CreateDirectory(Path.GetDirectoryName(shapefilePath));

                if (!File.Exists(shapefilePath))
                {
                    Log.recordMessage($"Shapefile not found at '{shapefilePath}'. Creating it...", Bis_Log_Message_Type.Note);
                    try
                    {
                        var njStatePlane = SpatialReferenceBuilder.CreateSpatialReference(requiredWkid);
                        var parameters = Geoprocessing.MakeValueArray(Path.GetDirectoryName(shapefilePath), Path.GetFileName(shapefilePath), "POLYGON", "", "DISABLED", "DISABLED", njStatePlane);
                        var gpResult = Geoprocessing.ExecuteToolAsync("management.CreateFeatureclass", parameters).Result;
                        if (gpResult.IsFailed)
                        {
                            string allMessages = string.Join("\n", gpResult.Messages.Select(m => m.Text));
                            Log.recordError($"Failed to create shapefile. GP Messages: {allMessages}", null, nameof(EnsureManualAddLayerExistsAsync));
                            return;
                        }

                        parameters = Geoprocessing.MakeValueArray(shapefilePath, "id", "TEXT", "", "", 50);
                        gpResult = Geoprocessing.ExecuteToolAsync("management.AddField", parameters).Result;
                        if (gpResult.IsFailed)
                        {
                            string allMessages = string.Join("\n", gpResult.Messages.Select(m => m.Text));
                            Log.recordError($"Failed to add 'id' field. GP Messages: {allMessages}", null, nameof(EnsureManualAddLayerExistsAsync));
                            return;
                        }
                    }
                    catch (Exception ex)
                    {
                        Log.recordError("An exception occurred during geoprocessing.", ex, nameof(EnsureManualAddLayerExistsAsync));
                        return;
                    }
                }

                var layerParams = new LayerCreationParams(new Uri(shapefilePath)) { Name = layerName };
                _manualAddLayer = LayerFactory.Instance.CreateLayer<FeatureLayer>(layerParams, map);

                if (_manualAddLayer != null)
                {
                    Log.recordMessage($"Successfully added layer '{layerName}' to the map and obtained a reference.", Bis_Log_Message_Type.Note);
                }
                else
                {
                    Log.recordError($"Could not create or find the '{layerName}' layer after all checks.", null, nameof(EnsureManualAddLayerExistsAsync));
                }
            });
        }
        #region


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