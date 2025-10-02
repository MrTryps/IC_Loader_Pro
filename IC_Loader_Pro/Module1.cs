using ArcGIS.Desktop.Core.Events;
using ArcGIS.Desktop.Framework;
using ArcGIS.Desktop.Framework.Contracts;
using ArcGIS.Desktop.Framework.Threading.Tasks;
using ArcGIS.Desktop.Mapping;
using IC_Rules_2025;
using System;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading.Tasks;


namespace IC_Loader_Pro
{
    internal class Module1 : ArcGIS.Desktop.Framework.Contracts.Module
    {
        private static Module1 _this = null;
        private static BIS_Log _log;
        private static IC_Rules _icRules = null;
        private static BisDbNpgsql _postGreTool = null;
        private static BisDbNjems _njemsTool;
        private static BisDbCompass _compassTool;
        private static BisDbAccess _accessTool;
        private static Bis_Regex _regexTool = null;
        private static BisFileTools _fileTool = null;

        #region Public Static Properties
        public static Module1 Current => _this;
        public static BIS_Log Log => _log;
        public static IC_Rules IcRules => _icRules;
        public static BisDbNpgsql PostGreTool => _postGreTool;
        public static BisDbNjems NjemsTool => _njemsTool;
        public static BisDbCompass CompassTool => _compassTool;
        public static BisDbAccess AccessTool => _accessTool;
        public static Bis_Regex RegexTool => _regexTool;
        public static BisFileTools FileTool => _fileTool;
        public static DateTime BuildDate { get; private set; }
        private static bool _initializationFailed = false;

        /// <summary>
        /// The required Well-Known ID (WKID) for the project's coordinate system.
        /// </summary>
        public const int RequiredWkid = 3424;

#if DEBUG
        public static bool IsInTestMode { get; set; } = true;
        #else
            public static bool IsInTestMode { get; set; } = false;
        #endif
        #endregion

        #region Overrides
        /// <summary>
        /// Called by the Framework when the Add-in is loaded.
        /// This is the ideal place to perform one-time initialization of any
        /// core services or resources used by your add-in.
        /// </summary>
        /// <returns>A Task that represents the initialization process.</returns>
        protected override bool Initialize()
        {
            if (_initializationFailed) return false;
            _this = this;        

            try
            {
                var attribute = Assembly.GetExecutingAssembly()
                                    .GetCustomAttributes<AssemblyMetadataAttribute>()
                                    .FirstOrDefault(a => a.Key == "BuildDate");
                if (attribute != null)
                {
                    BuildDate = DateTime.Parse(attribute.Value, CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind);
                }
            }
            catch (Exception ex)
            {
                BuildDate = DateTime.MinValue;
            }


            try
            {
                _log = new BIS_Log("IC_Loader_Pro");
                if (BuildDate != DateTime.MinValue)
                {
                    _log.RecordMessage($"IC Loader Pro Build Date: {BuildDate.ToLocalTime():yyyy-MM-dd HH:mm:ss}", BIS_Log.BisLogMessageType.Note);
                    _log.AddBlankLine();
                }
                _fileTool = new BisFileTools(_log);
                _regexTool = new Bis_Regex(_log);
                _postGreTool = new BisDbNpgsql(_log);
                _njemsTool = new BisDbNjems(_log);
                _compassTool = new BisDbCompass(_log);
                _accessTool = new BisDbAccess(_log);
                _icRules = new IC_Rules(_log, _postGreTool, _compassTool, _njemsTool, _accessTool, _fileTool, _regexTool);
                CleanupOrphanedTempFolders();
            }
            catch (Exception ex)
            {
                _initializationFailed = true;
                // Create a temporary, failsafe logger just in case the main one failed to initialize.
                var tempLog = _log ?? new BIS_Log("IC_Loader_Pro_Failsafe");
                string errorMessage = "A critical error occurred during IC Loader Pro initialization (likely a missing database driver). The add-in will be disabled for this session.";

                // Log the detailed error for troubleshooting.
                tempLog.RecordError($"FATAL: {errorMessage}", ex, nameof(Initialize));

                // Show a single, clear message to the user.
                ArcGIS.Desktop.Framework.Dialogs.MessageBox.Show(errorMessage, "IC Loader Pro - Initialization Error");

                // Returning false tells ArcGIS Pro to gracefully disable the module without crashing.
                return false;
            }

            try
            {
                var attribute = Assembly.GetExecutingAssembly()
                                    .GetCustomAttributes<AssemblyMetadataAttribute>()
                                    .FirstOrDefault(a => a.Key == "BuildDate");
                if (attribute != null)
                {
                    // Parse the ISO 8601 date format from the attribute
                    BuildDate = DateTime.Parse(attribute.Value, CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind);
                }
            }
            catch (Exception ex)
            {
                // Fallback in case of an error
                BuildDate = DateTime.MinValue;
                Log?.RecordError("Could not parse build date from assembly.", ex, "Initialize");
            }
            
            ProjectClosingEvent.Subscribe(OnProjectClosing);
            return true;
        }

        /// <summary>
        /// Called by Framework when ArcGIS Pro is closing
        /// </summary>
        /// <returns>False to prevent Pro from closing, otherwise True</returns>
        protected override bool CanUnload()
        {
            // Add any cleanup logic here if needed.
            return true;
        }

        /// <summary>
        /// Called just before the project closes.
        /// </summary>
        private Task OnProjectClosing(ProjectClosingEventArgs arg)
        {
            // All map modifications must be run on the main thread via QueuedTask.
            return QueuedTask.Run(() =>
            {
                // Get the active map if it exists.
                var activeMap = MapView.Active?.Map;
                if (activeMap == null)
                {
                    return; // No active map, nothing to clear.
                }

                Log.RecordMessage("Project closing. Clearing graphics layers...", BIS_Log.BisLogMessageType.Note);

                // Find and clear the main shapes layer
                var drawLayer = activeMap.FindLayers("IC Loader Shapes").FirstOrDefault() as GraphicsLayer;
                if (drawLayer != null)
                {
                    drawLayer.RemoveElements();
                }

                // Find and clear the highlight layer
                var highlightLayer = activeMap.FindLayers("IC Loader Highlight").FirstOrDefault() as GraphicsLayer;
                if (highlightLayer != null)
                {
                    highlightLayer.RemoveElements();
                }
            });
        }

        private void CleanupOrphanedTempFolders()
        {
            try
            {
                string localAppData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
                string addinTempRoot = Path.Combine(localAppData, "IC_Loader_Pro_Temp");

                if (Directory.Exists(addinTempRoot))
                {
                    Log.RecordMessage("Performing startup cleanup of temporary files...", BIS_Log.BisLogMessageType.Note);
                    // Delete all subdirectories (the GUID folders) but leave the root folder.
                    foreach (var directory in Directory.GetDirectories(addinTempRoot))
                    {
                        Directory.Delete(directory, true);
                    }
                }
            }
            catch (Exception ex)
            {
                // Log the error but don't prevent the add-in from loading.
                Log.RecordError("An error occurred during startup cleanup of temporary folders.", ex, "CleanupOrphanedTempFolders");
            }
        }
        #endregion Overrides

    }
}
