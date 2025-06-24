using ArcGIS.Core.CIM;
using ArcGIS.Core.Geometry;
using ArcGIS.Desktop.Core;
using ArcGIS.Desktop.Core.Geoprocessing;
using ArcGIS.Desktop.Framework;
using ArcGIS.Desktop.Framework.Threading.Tasks;
using ArcGIS.Desktop.Mapping;
using ArcGIS.Desktop.Mapping.Events;
using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using static BIS_Tools_2025_Core.BIS_Log;
using static IC_Loader_Pro.Module1;

namespace IC_Loader_Pro
{
    // Note the 'partial' keyword. This merges this file with your main ViewModel file.
    internal partial class Dockpane_IC_LoaderViewModel
    {
        #region Initialization Fields and Properties

        private FeatureLayer _manualAddLayer = null;

        /// <summary>
        /// A helper property that returns the full, persistent path for our 'manually_added' shapefile.
        /// </summary>
        private string ManualAddShapefilePath
        {
            get
            {
                string localAppDataFolder = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
                string appFolderPath = Path.Combine(localAppDataFolder, "IC_Loader_Pro");
                Directory.CreateDirectory(appFolderPath);
                return Path.Combine(appFolderPath, "manually_added.shp");
            }
        }

        #endregion

        #region Initialization Methods

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

        /// <summary>
        /// This is called by the framework when the dockpane is first created.
        /// Perfect for one-time setup like subscribing to events or loading initial data.
        /// </summary>
        protected override Task InitializeAsync()
        {
            ActiveMapViewChangedEvent.Subscribe(OnActiveMapViewChanged);
            if (MapView.Active != null)
            {
                OnActiveMapViewChanged(new ActiveMapViewChangedEventArgs(MapView.Active, null));
            }
            return Task.CompletedTask;
        }

        /// <summary>
        /// Contains all the main setup logic for the dockpane.
        /// It is given a valid map to work with.
        /// </summary>
        private async Task LoadAndInitializeAsync(Map activeMap)
        {
            // Use a lock and a flag to ensure this complex initialization only ever runs once
            // for the lifetime of the dockpane.
            lock (_lock)
            {
                if (_isInitialized) return;
                _isInitialized = true;
            }

            StatusMessage = "Initializing...";
            IsUIEnabled = false;

            try
            {
                // Now, we just need to ensure the SR and the layer are correct on the provided map.
                await QueuedTask.Run(() =>
                {
                    int requiredWkid = 2260; // NAD 1983 StatePlane New Jersey FIPS 2900 (US Feet)
                    if (activeMap.SpatialReference?.Wkid != requiredWkid)
                    {
                        Log.recordMessage($"Active map is not in the required coordinate system. Forcing it to NJ State Plane.", Bis_Log_Message_Type.Warning);
                        var njStatePlane = SpatialReferenceBuilder.CreateSpatialReference(requiredWkid);
                        activeMap.SetSpatialReference(njStatePlane);
                    }
                });

                // Ensure our special "manually_added" scratch layer is ready.
                await EnsureManualAddLayerExistsAsync(activeMap);

                // Load the data for the IC Queue toggle buttons.
                Log.recordMessage("About to call [RefreshICQueuesAsync]", Bis_Log_Message_Type.Note);
                await RefreshICQueuesAsync();


                // Final step: Update status and enable the UI.
                StatusMessage = "Ready. Please select an IC Type.";
                IsUIEnabled = true;
            }
            catch (Exception ex)
            {
                Log.recordError("A fatal error occurred during initialization.", ex, nameof(LoadAndInitializeAsync));
                StatusMessage = "An error occurred during initialization.";
            }
        }

       

        /// <summary>
        /// Ensures the "manually_added" scratch layer exists and is valid.
        /// </summary>
        private Task EnsureManualAddLayerExistsAsync(Map map)
        {
            return QueuedTask.Run(() =>
            {
                // This is the full, robust method we built previously
                string layerName = "manually_added";
                int requiredWkid = 2260;

                var existingLayer = map.GetLayersAsFlattenedList().FirstOrDefault(l => l.Name.Equals(layerName, StringComparison.CurrentCultureIgnoreCase)) as FeatureLayer;
                if (existingLayer != null)
                {
                    bool isLayerValid = false;
                    string validationError = "Unknown validation error.";
                    try
                    {
                        using (var featureClass = existingLayer.GetFeatureClass())
                        {
                            if (featureClass == null) { validationError = "Data source is broken."; }
                            else
                            {
                                var definition = featureClass.GetDefinition();
                                if (definition.GetShapeType() != GeometryType.Polygon) { validationError = "Geometry type is not Polygon."; }
                                else if (!definition.GetFields().Any(f => f.Name.Equals("id", StringComparison.CurrentCultureIgnoreCase))) { validationError = "Required 'id' field is missing."; }
                                else if (definition.GetSpatialReference()?.Wkid != requiredWkid) { validationError = $"Incorrect coordinate system."; }
                                else { isLayerValid = true; }
                            }
                        }
                    }
                    catch (Exception ex) { validationError = $"Validation exception: {ex.Message}"; }

                    if (isLayerValid)
                    {
                        _manualAddLayer = existingLayer;
                        return;
                    }
                    else
                    {
                        Log.recordMessage($"Removing invalid '{layerName}' layer. Reason: {validationError}", Bis_Log_Message_Type.Warning);
                        map.RemoveLayer(existingLayer);
                    }
                }

                string shapefilePath = this.ManualAddShapefilePath;
                if (!File.Exists(shapefilePath))
                {
                    try
                    {
                        var njStatePlane = SpatialReferenceBuilder.CreateSpatialReference(requiredWkid);
                        var parameters = Geoprocessing.MakeValueArray(Path.GetDirectoryName(shapefilePath), Path.GetFileName(shapefilePath), "POLYGON", "", "DISABLED", "DISABLED", njStatePlane);
                        var gpResult = Geoprocessing.ExecuteToolAsync("management.CreateFeatureclass", parameters).Result;
                        if (gpResult.IsFailed) { Log.recordError($"Failed to create shapefile: {string.Join("\n", gpResult.Messages.Select(m => m.Text))}", null, nameof(EnsureManualAddLayerExistsAsync)); return; }

                        parameters = Geoprocessing.MakeValueArray(shapefilePath, "id", "TEXT", "", "", 50);
                        gpResult = Geoprocessing.ExecuteToolAsync("management.AddField", parameters).Result;
                        if (gpResult.IsFailed) { Log.recordError($"Failed to add 'id' field: {string.Join("\n", gpResult.Messages.Select(m => m.Text))}", null, nameof(EnsureManualAddLayerExistsAsync)); return; }
                    }
                    catch (Exception ex) { Log.recordError("Exception during geoprocessing.", ex, nameof(EnsureManualAddLayerExistsAsync)); return; }
                }

                var layerParams = new LayerCreationParams(new Uri(shapefilePath)) { Name = layerName };
                _manualAddLayer = LayerFactory.Instance.CreateLayer<FeatureLayer>(layerParams, map);
                if (_manualAddLayer == null) { Log.recordError($"Could not create or find the '{layerName}' layer after all checks.", null, nameof(EnsureManualAddLayerExistsAsync)); }
            });
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

        #endregion
    }
}