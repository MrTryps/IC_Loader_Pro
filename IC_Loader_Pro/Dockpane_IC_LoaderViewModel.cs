using ArcGIS.Core.CIM;
using ArcGIS.Core.Geometry;
using ArcGIS.Desktop.Core;
using ArcGIS.Desktop.Core.Geoprocessing;
using ArcGIS.Desktop.Framework;
using ArcGIS.Desktop.Framework.Contracts;
using ArcGIS.Desktop.Framework.Threading.Tasks;
using ArcGIS.Desktop.Mapping;
using IC_Loader_Pro.Models;
using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Controls;
using System.Windows.Input;
using static BIS_Tools_2025_Core.BIS_Log;
using static IC_Loader_Pro.Module1;

namespace IC_Loader_Pro
{
    internal partial class Dockpane_IC_LoaderViewModel : DockPane
    {
        private bool _isInitialized = false;
        private FeatureLayer _manualAddLayer = null;
        private readonly object _lock = new object();
        public const string DockPaneId = "IC_Loader_Pro_Dockpane_IC_Loader";

        #region UI Properties

        private bool _isUIEnabled = false;
        /// <summary>
        /// Controls whether the main UI controls are enabled.
        /// The buttons' CanExecute condition is bound to this property.
        /// </summary>
        public bool IsUIEnabled
        {
            get => _isUIEnabled;
            set => SetProperty(ref _isUIEnabled, value);
        }

        private string _statusMessage = "Please open or create a project to begin.";
        /// <summary>
        /// A message displayed to the user at the top of the dockpane.
        /// </summary>
        public string StatusMessage
        {
            get => _statusMessage;
            set => SetProperty(ref _statusMessage, value);
        }

        // ... your other UI properties like ICQueues, SelectedQueue, etc. go here ...

        #endregion

        internal static void Show()
        {
            // Use the DockPaneManager to find the pane. This is the correct method.
            var pane = FrameworkApplication.DockPaneManager.Find(DockPaneId);
            if (pane == null)
                return;

            pane.Activate();
        }



        /// <summary>
        /// This is the single entry point for our setup logic.
        /// It will be called from the ribbon button AFTER the dockpane is shown.
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

            SaveCommand = new RelayCommand(() => OnSave(), () => IsUIEnabled);
            SkipCommand = new RelayCommand(() => OnSkip(), () => IsUIEnabled);
            RejectCommand = new RelayCommand(() => OnReject(), () => IsUIEnabled);

            Log.recordMessage("Initializing Dockpane...", Bis_Log_Message_Type.Note);

            try
            {
                // We can now safely assume a map exists and is active.
                Map activeMap = MapView.Active?.Map;
                if (activeMap == null)
                {
                    Log.recordError("No active map found. Initialization cannot proceed.", null, nameof(LoadAndInitializeAsync));
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

        /// <summary>
        /// Ensures the "manually_added" scratch layer exists, checking the map, then disk, 
        /// and finally creating it if it doesn't exist or is invalid.
        /// </summary>
        /// <param name="map">The map to add the layer to.</param>
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
    }
}