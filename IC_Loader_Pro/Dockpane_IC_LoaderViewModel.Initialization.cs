using ArcGIS.Core.CIM;
using ArcGIS.Core.Geometry;
using ArcGIS.Desktop.Core;
using ArcGIS.Desktop.Core.Geoprocessing;
using ArcGIS.Desktop.Framework;
using ArcGIS.Desktop.Framework.Threading.Tasks;
using ArcGIS.Desktop.Mapping;
using BIS_Tools_DataModels_2025;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using static BIS_Log;
using static IC_Loader_Pro.Module1;

namespace IC_Loader_Pro
{
    // Note the 'partial' keyword. This merges this file with your main ViewModel file.
    internal partial class Dockpane_IC_LoaderViewModel
    {
        #region Initialization Fields and Properties

        private FeatureLayer _manualAddLayer = null;
        private bool _isUIEnabled = false;
        public bool IsUIEnabled { get => _isUIEnabled; set => SetProperty(ref _isUIEnabled, value); }


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
        /// This is the single entry point for our setup logic, called by the button click.
        /// </summary>
        public async Task LoadAndInitializeAsync()
        {
            // Use a lock and a flag to ensure this complex initialization only ever runs once.
            lock (_lock)
            {
                if (_isInitialized)
                    return;
                _isInitialized = true;
            }

            Log.RecordMessage("Initializing Dockpane...", BisLogMessageType.Note);
            StatusMessage = "Initializing map and layers...";
            IsUIEnabled = false; // Disable UI during setup

            try
            {
                Map activeMap = await GetAndPrepareMapAsync();
                if (activeMap == null)
                {
                    StatusMessage = "Error: A map could not be opened or created.";
                    return;
                }

                await EnsureManualAddLayerExistsAsync(activeMap);
                await EnsureGraphicsLayersExistsAsync(activeMap);

                Log.RecordMessage("Refreshing IC Queues...", BisLogMessageType.Note);
                await RefreshICQueuesAsync();

                Log.RecordMessage("Initialization complete.", BisLogMessageType.Note);
                Log.AddBlankLine();
                StatusMessage = "Ready. Please select an IC Type.";
                IsUIEnabled = true; // Enable the UI now that setup is complete
            }
            catch (Exception ex)
            {
                Log.RecordError("A fatal error occurred during map and layer initialization.", ex, nameof(LoadAndInitializeAsync));
                StatusMessage = "An error occurred during initialization.";
            }
        }

        /// <summary>
        /// Gets an active map, or creates a new one if necessary, and ensures it
        /// is set to the required NJ State Plane coordinate system.
        /// </summary>
        private async Task<Map> GetAndPrepareMapAsync()
        {
            if (Project.Current == null)
            {
                Log.RecordError("Cannot get or create a map because no project is open.", null, nameof(GetAndPrepareMapAsync));
                return null;
            }

            try
            {
                Map map = MapView.Active?.Map;

                if (map == null)
                {
                    var mapProjectItem = Project.Current.GetItems<MapProjectItem>().FirstOrDefault();
                    if (mapProjectItem != null)
                    {
                        map = await QueuedTask.Run(() => mapProjectItem.GetMap());
                        await ProApp.Panes.CreateMapPaneAsync(map);
                    }
                }

                if (map == null)
                {
                    await QueuedTask.Run(() =>
                    {
                        // --- THE CORRECTED LOGIC ---
                        // Use a standard if-check for the basemap instead of the '??' operator.
                        Basemap basemap = Basemap.ProjectDefault;
                        if (basemap == null)
                        {
                            Log.RecordMessage("No default basemap found. Falling back to 'Streets'.", BisLogMessageType.Note);
                            basemap = Basemap.Streets;
                        }

                        map = MapFactory.Instance.CreateMap("New Map", MapType.Map, MapViewingMode.Map, basemap);
                    });
                    await ProApp.Panes.CreateMapPaneAsync(map);
                }

                // Now that we have a map, ENSURE it has the correct spatial reference.
                await QueuedTask.Run(() =>
                {
                    int requiredWkid = 2260; // NAD 1983 StatePlane New Jersey FIPS 2900 (US Feet)
                    if (map.SpatialReference?.Wkid != requiredWkid)
                    {
                        Log.RecordMessage($"Map '{map.Name}' is not in the required coordinate system. Forcing it to NJ State Plane.", BisLogMessageType.Warning);
                        var njStatePlane = SpatialReferenceBuilder.CreateSpatialReference(requiredWkid);
                        map.SetSpatialReference(njStatePlane);
                    }
                });

                return map;
            }
            catch (Exception ex)
            {
                Log.RecordError("An unexpected error occurred while getting or preparing the map.", ex, nameof(GetAndPrepareMapAsync));
                return null;
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
                        Log.RecordMessage($"Removing invalid '{layerName}' layer. Reason: {validationError}", BisLogMessageType.Warning);
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
                        if (gpResult.IsFailed) { Log.RecordError($"Failed to create shapefile: {string.Join("\n", gpResult.Messages.Select(m => m.Text))}", null, nameof(EnsureManualAddLayerExistsAsync)); return; }

                        parameters = Geoprocessing.MakeValueArray(shapefilePath, "id", "TEXT", "", "", 50);
                        gpResult = Geoprocessing.ExecuteToolAsync("management.AddField", parameters).Result;
                        if (gpResult.IsFailed) { Log.RecordError($"Failed to add 'id' field: {string.Join("\n", gpResult.Messages.Select(m => m.Text))}", null, nameof(EnsureManualAddLayerExistsAsync)); return; }
                    }
                    catch (Exception ex) { Log.RecordError("Exception during geoprocessing.", ex, nameof(EnsureManualAddLayerExistsAsync)); return; }
                }

                var layerParams = new LayerCreationParams(new Uri(shapefilePath)) { Name = layerName };
                _manualAddLayer = LayerFactory.Instance.CreateLayer<FeatureLayer>(layerParams, map);
                if (_manualAddLayer == null) { Log.RecordError($"Could not create or find the '{layerName}' layer after all checks.", null, nameof(EnsureManualAddLayerExistsAsync)); }
            });
        }
        /// <summary>
        /// Ensures a dedicated graphics layer for displaying IC shapes exists on the map.
        /// </summary>
        private Task EnsureGraphicsLayersExistsAsync(Map map)
        {
            const string drawLayerName = "IC Loader Shapes";
            const string highlightLayerName = "IC Loader Highlight";

            return QueuedTask.Run(() =>
            {
                // Check if the layer already exists
                var drawLayer = map.FindLayers(drawLayerName).FirstOrDefault() as GraphicsLayer;
                if (drawLayer != null)
                {
                    drawLayer.RemoveElements();
                   // Log.RecordMessage($"Graphics layer '{drawLayerName}' already exists.", BisLogMessageType.Note);                   
                }
                else // If not, create it and add it to the top of the map
                {
                    Log.RecordMessage($"Creating new graphics layer: '{drawLayerName}'", BisLogMessageType.Note);
                    GraphicsLayerCreationParams graphicsLayerCreationParams = new GraphicsLayerCreationParams
                    {
                        Name = drawLayerName
                    };
                    drawLayer = LayerFactory.Instance.CreateLayer<GraphicsLayer>(graphicsLayerCreationParams, map);
                }

                _graphicsLayer = drawLayer;

                var drawLayerDef = drawLayer.GetDefinition() as CIMGraphicsLayer;
                if (drawLayerDef != null)
                {
                    // 2. Modify the blueprint
                    drawLayerDef.Selectable = true;
                    // 3. Apply the modified blueprint back to the layer
                    drawLayer.SetDefinition(drawLayerDef);
                }

                // Check for the highlight layer
                var highlightLayer = map.FindLayers(highlightLayerName).FirstOrDefault() as GraphicsLayer;
                if (highlightLayer != null)
                {
                    highlightLayer.RemoveElements(); // Clear existing elements
                   // Log.RecordMessage($"Highlight layer '{highlightLayerName}' already exists.", BisLogMessageType.Note);
                }
                else
                {
                    // Create the highlight layer and add it to the top of the map
                    GraphicsLayerCreationParams graphicsLayerCreationParams = new GraphicsLayerCreationParams
                    {
                        Name = highlightLayerName
                    };
                    GraphicsLayer newGraphicsLayer = LayerFactory.Instance.CreateLayer<GraphicsLayer>(graphicsLayerCreationParams, map);
                }           
            });
        }


        #endregion
    }
}