using ArcGIS.Core.CIM;
using ArcGIS.Core.Geometry; // Required for SpatialReferences
using ArcGIS.Desktop.Core;
using ArcGIS.Desktop.Core.Geoprocessing; // Required for Geoprocessing
using ArcGIS.Desktop.Framework;
using ArcGIS.Desktop.Framework.Contracts;
using ArcGIS.Desktop.Framework.Dialogs;
using ArcGIS.Desktop.Framework.Threading.Tasks;
using ArcGIS.Desktop.Mapping;
using BIS_Tools_2025_Core;
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


namespace IC_Loader_Pro
{
    // This class is now both the DockPane and the ViewModel
    internal class Dockpane_IC_LoaderViewModel : DockPane
    {
        private const string _dockPaneID = "IC_Loader_Pro_Dockpane_IC_Loader";

        #region Properties
        public ObservableCollection<Models.ICQueueInfo> ICQueues { get; } = new ObservableCollection<Models.ICQueueInfo>();
        private FeatureLayer _manualAddLayer = null;

        private Models.ICQueueInfo _selectedQueue;
        public Models.ICQueueInfo SelectedQueue
        {
            get => _selectedQueue;
            set => SetProperty(ref _selectedQueue, value, () => SelectedQueue);
        }

        private Models.ActiveEmail _currentEmail;
        public Models.ActiveEmail CurrentEmail
        {
            get => _currentEmail;
            set => SetProperty(ref _currentEmail, value, () => CurrentEmail);
        }

        // Aggregate Counts
        private int _passedCount;
        public int PassedCount { get => _passedCount; set => SetProperty(ref _passedCount, value, () => PassedCount); }
        private int _skippedCount;
        public int SkippedCount { get => _skippedCount; set => SetProperty(ref _skippedCount, value, () => SkippedCount); }
        private int _failedCount;
        public int FailedCount { get => _failedCount; set => SetProperty(ref _failedCount, value, () => FailedCount); }
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

            LoadDummyData();        
        }


        /// <summary>
        /// The main entry point to ensure a map is ready for the add-in.
        /// Creates a map if one does not exist, then ensures the scratch layer is present.
        /// </summary>
        private async Task InitializeMapAndLayersAsync()
        {
            try
            {
                Map activeMap = MapView.Active?.Map;

                if (activeMap == null)
                {
                    await QueuedTask.Run(() =>
                    {
                        activeMap = MapFactory.Instance.CreateMap("New Map", MapType.Map);
                    });
                    await ProApp.Panes.CreateMapPaneAsync(activeMap);
                }

                // Now that a map is guaranteed to exist, ensure our special "manually_added" scratch layer is ready.
                await EnsureManualAddLayerExistsAsync(activeMap);
            }
            catch (Exception ex)
            {
                Module1._Log.recordError("A fatal error occurred during map and layer initialization.", ex, nameof(InitializeMapAndLayersAsync));
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

                if (map.GetLayersAsFlattenedList().FirstOrDefault(l => l.Name.Equals(layerName, StringComparison.CurrentCultureIgnoreCase)) is FeatureLayer existingLayer)
                {
                    Module1._Log.recordMessage($"Layer '{layerName}' already exists in the map.", Bis_Log_Message_Type.Note);
                    _manualAddLayer = existingLayer;
                    return;
                }

                string tempFolder = Path.GetTempPath();
                string shapefilePath = Path.Combine(tempFolder, layerName + ".shp");

                if (!File.Exists(shapefilePath))
                {
                    Module1._Log.recordMessage($"Shapefile not found at '{shapefilePath}'. Creating it...", Bis_Log_Message_Type.Note);
                    try
                    {
                        var parameters = Geoprocessing.MakeValueArray(tempFolder, layerName, "POLYGON", "", "DISABLED", "DISABLED", SpatialReferences.WGS84);
                        var gpResult = Geoprocessing.ExecuteToolAsync("management.CreateFeatureclass", parameters).Result;

                        if (gpResult.IsFailed)
                        {
                            Module1._Log.recordError($"Failed to create shapefile. GP Messages: {gpResult.Messages}", null, nameof(EnsureManualAddLayerExistsAsync));
                            return;
                        }

                        parameters = Geoprocessing.MakeValueArray(shapefilePath, "id", "TEXT", "", "", 50);
                        gpResult = Geoprocessing.ExecuteToolAsync("management.AddField", parameters).Result;

                        if (gpResult.IsFailed)
                        {
                            Module1._Log.recordError($"Failed to add 'id' field. GP Messages: {gpResult.Messages}", null, nameof(EnsureManualAddLayerExistsAsync));
                            return;
                        }
                    }
                    catch (Exception ex)
                    {
                        Module1._Log.recordError("An exception occurred during geoprocessing.", ex, nameof(EnsureManualAddLayerExistsAsync));
                        return;
                    }
                }

                var layerParams = new LayerCreationParams(new Uri(shapefilePath)) { Name = layerName };
                _manualAddLayer = LayerFactory.Instance.CreateLayer<FeatureLayer>(layerParams, map);

                if (_manualAddLayer != null)
                {
                    Module1._Log.recordMessage($"Successfully added layer '{layerName}' to the map and obtained a reference.", Bis_Log_Message_Type.Note);
                }
                else
                {
                    Module1._Log.recordError($"Could not create or find the '{layerName}' layer after all checks.", null, nameof(EnsureManualAddLayerExistsAsync));
                }
            });
        }                 


        private void LoadDummyData()
        {
            // It's good practice to clear the list first to prevent duplicates
            // if this method were ever called more than once.
            ICQueues.Clear();

            // Create and add some sample ICQueueInfo objects to the collection.
            // The UI will create one button for each item added here.
            ICQueues.Add(new Models.ICQueueInfo { Name = "CEAs", EmailCount = 12 });
            ICQueues.Add(new Models.ICQueueInfo { Name = "DNAs", EmailCount = 5 });
            ICQueues.Add(new Models.ICQueueInfo { Name = "WRAs", EmailCount = 21 });            

            // Set a default selection so the first button is active when the pane opens.
            SelectedQueue = ICQueues.FirstOrDefault();

            // You can also populate the other dummy data here if it's not already
            PassedCount = 11;
            SkippedCount = 2;
            FailedCount = 10;

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