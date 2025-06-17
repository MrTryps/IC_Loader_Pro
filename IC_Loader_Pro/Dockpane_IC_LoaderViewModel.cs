using ArcGIS.Core.CIM;
using ArcGIS.Core.Geometry; // Required for SpatialReferences
using ArcGIS.Desktop.Core;
using ArcGIS.Desktop.Core.Geoprocessing; // Required for Geoprocessing
using ArcGIS.Desktop.Framework;
using ArcGIS.Desktop.Framework.Contracts;
using ArcGIS.Desktop.Framework.Dialogs;
using ArcGIS.Desktop.Framework.Threading.Tasks;
using ArcGIS.Desktop.Layouts;
using ArcGIS.Desktop.Mapping;
using BIS_Tools_2025_Core;
using IC_Loader_Pro.Models;
using IC_Rules_2025;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO; // Required for Path
using System.Linq;
using System.Reflection;
using System.Threading.Tasks;
using System.Windows.Input;
using static BIS_Tools_2025_Core.BIS_Log;
using static IC_Loader_Pro.Module1;


namespace IC_Loader_Pro
{
    // This class is now both the DockPane and the ViewModel
    internal class Dockpane_IC_LoaderViewModel : DockPane
    {
        private const string _dockPaneID = "IC_Loader_Pro_Dockpane_IC_Loader";

        #region Properties
        public ObservableCollection<ICQueueInfo> ICQueues { get; } = new ObservableCollection<ICQueueInfo>();
        private FeatureLayer _manualAddLayer = null;

        private ICQueueInfo _selectedQueue;
        public ICQueueInfo SelectedQueue
        {
            get => _selectedQueue;
            set => SetProperty(ref _selectedQueue, value, () => SelectedQueue);
        }

        private ActiveEmail _currentEmail;
        public ActiveEmail CurrentEmail
        {
            get => _currentEmail;
            set => SetProperty(ref _currentEmail, value, () => CurrentEmail);
        }
        #endregion

        #region Commands
        public ICommand ShowNotesCommand { get; }
        public ICommand SearchCommand { get; }
        public ICommand ToolsCommand { get; }
        public ICommand OptionsCommand { get; }
        #endregion

        /// <summary>
        /// The constructor is now responsible for creating the View
        /// and setting it as the dock pane's content.
        /// </summary>
        public Dockpane_IC_LoaderViewModel()
        {
            // Create an instance of the UserControl and set it as the Content.
            // The UserControl will automatically inherit this class instance as its DataContext.
            this.Content = new Dockpane_IC_LoaderView();

            // Initialize commands
            ShowNotesCommand = new RelayCommand(() => MessageBox.Show("Show Notes command executed."));
            SearchCommand = new RelayCommand(() => MessageBox.Show("Search command executed."));
            ToolsCommand = new RelayCommand(() => MessageBox.Show("Tools command executed."));
            OptionsCommand = new RelayCommand(() => MessageBox.Show("Options command executed."));

            _ = InitializeMapAndLayersAsync();

            LoadSampleICQueueData();        
        }

        /// <summary>
        /// A helper property that returns the full, persistent path for our shapefile.
        /// </summary>
        private string ManualAddShapefilePath
        {
            get
            {
                // Get the path to the user's local app data folder. This is a safe, persistent place to write files.
                string localAppDataFolder = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
                // Define a specific folder for your application to keep things tidy.
                string appFolderPath = Path.Combine(localAppDataFolder, "IC_Loader_Pro");
                // Ensure this directory exists before you use it.
                Directory.CreateDirectory(appFolderPath);
                // Return the full path to the shapefile
                return Path.Combine(appFolderPath, "manually_added.shp");
            }
        }



        private async Task InitializeMapAndLayersAsync()
        {
            try
            {
                Map activeMap = null;

                if (MapView.Active != null)
                {
                    activeMap = MapView.Active.Map;
                }

                if (activeMap == null)
                {
                    var mapProjectItem = Project.Current.GetItems<MapProjectItem>().FirstOrDefault();
                    if (mapProjectItem != null)
                    {
                        Log.recordMessage("Active map view not found. Opening first map found in project.", Bis_Log_Message_Type.Note);
                        activeMap = await QueuedTask.Run(() => mapProjectItem.GetMap());
                        await ProApp.Panes.CreateMapPaneAsync(activeMap);
                    }
                }

                if (activeMap == null)
                {
                    Log.recordMessage("No maps found in project. Creating a new map.", Bis_Log_Message_Type.Note);
                    await QueuedTask.Run(() =>
                    {
                        // --- THIS IS THE CORRECTED LOGIC ---
                        // Use a standard if-check for the basemap.
                        Basemap basemap = Basemap.ProjectDefault;
                        if (basemap == null)
                        {
                            Log.recordMessage("No default basemap found. Falling back to 'Streets'.", Bis_Log_Message_Type.Note);
                            basemap = Basemap.Streets;
                        }

                        activeMap = MapFactory.Instance.CreateMap("New Map", MapType.Map, MapViewingMode.Map, basemap);
                        var njStatePlane = SpatialReferenceBuilder.CreateSpatialReference(2260);
                        activeMap.SetSpatialReference(njStatePlane);
                    });

                    await ProApp.Panes.CreateMapPaneAsync(activeMap);
                }

                // Now that a map is guaranteed to exist, we run our final checks and setup on it.
                await QueuedTask.Run(() =>
                {
                    int requiredWkid = 2260;
                    if (activeMap.SpatialReference?.Wkid != requiredWkid)
                    {
                        Log.recordMessage($"Active map is not in the required coordinate system. Forcing it to NJ State Plane (WKID {requiredWkid}).", Bis_Log_Message_Type.Warning);
                        var njStatePlane = SpatialReferenceBuilder.CreateSpatialReference(requiredWkid);
                        activeMap.SetSpatialReference(njStatePlane);
                    }
                });

                await EnsureManualAddLayerExistsAsync(activeMap);
            }
            catch (Exception ex)
            {
                Log.recordError("A fatal error occurred during map and layer initialization.", ex, nameof(InitializeMapAndLayersAsync));
            }
        }


        /// <summary>
        /// Ensures the "manually_added" scratch layer exists, checking the map, then disk, and creating it if needed.
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
                        // Getting the feature class can throw an exception if the link is badly broken.
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
                        // The layer is invalid. Log the reason and remove it so we can create a fresh one.
                        Log.recordMessage($"Removing invalid '{layerName}' layer. Reason: {validationError}", Bis_Log_Message_Type.Warning);
                        map.RemoveLayer(existingLayer);
                    }
                }

                // --- The rest of the method continues as before ---
                // (If we get to this point, it means the layer did not exist or was invalid and has been removed)

                string shapefilePath = this.ManualAddShapefilePath;

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
                            Log.recordError($"Failed to create shapefile. GP Messages: {gpResult.Messages}", null, nameof(EnsureManualAddLayerExistsAsync));
                            return;
                        }

                        parameters = Geoprocessing.MakeValueArray(shapefilePath, "id", "TEXT", "", "", 50);
                        gpResult = Geoprocessing.ExecuteToolAsync("management.AddField", parameters).Result;
                        if (gpResult.IsFailed)
                        {
                            Log.recordError($"Failed to add 'id' field. GP Messages: {gpResult.Messages}", null, nameof(EnsureManualAddLayerExistsAsync));
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


        private void LoadSampleICQueueData()
        {
            // It's good practice to clear the list first to prevent duplicates
            // if this method were ever called more than once.
            ICQueues.Clear();

            // Create and add some sample ICQueueInfo objects to the collection.
            // The UI will create one button for each item added here.
            ICQueues.Add(new Models.ICQueueInfo { Name = "CEAs", EmailCount = 12, PassedCount = 0, SkippedCount = 0, FailedCount = 0 });
            ICQueues.Add(new Models.ICQueueInfo { Name = "DNAs", EmailCount = 5, PassedCount = 0, SkippedCount = 0, FailedCount = 0 });
            ICQueues.Add(new Models.ICQueueInfo { Name = "WRAs", EmailCount = 21, PassedCount = 0, SkippedCount = 0, FailedCount = 0 });            

            // Set a default selection so the first button is active when the pane opens.
            SelectedQueue = ICQueues.FirstOrDefault();

            CurrentEmail = new Models.ActiveEmail
            {
                Subject = "FW: Institutional Control Submission - Site 123",
                PrefID = "g0000355",
                DelID = "gIS_1234"
            };
        }

        /// <summary>
        /// Show the DockPane.
        /// </summary>
        internal static void Show()
        {
            DockPane pane = FrameworkApplication.DockPaneManager.Find(_dockPaneID);
            if (pane == null)
                return;
            pane.Activate();
        }
    }
}