using ArcGIS.Desktop.Core.Events;
using ArcGIS.Desktop.Framework;
using ArcGIS.Desktop.Framework.Contracts;
using ArcGIS.Desktop.Framework.Threading.Tasks;
using ArcGIS.Desktop.Mapping;
using IC_Rules_2025;
using System;
using System.Linq;
using System.Threading.Tasks;


namespace IC_Loader_Pro
{
    internal class Module1 : Module
    {
        private static Module1 _this = null;
        private static BIS_Log _log;
        private static IC_Rules _icRules = null;
        private static BIS_DB_PostGre _postGreTool = null;
        private static BisDbNjems _njemsTool;
        private static BisDbCompass _compassTool;
        private static BisDbAccess _accessTool;
        private static Bis_Regex _regexTool = null;
        private static BisFileTools _fileTool = null;

        /// <summary>
        /// Retrieve the singleton instance to this module here
        /// </summary>
        /// 

        #region Public Static Properties (Service Accessors)

        /// <summary>
        /// Retrieve the singleton instance of this module.
        /// </summary>
        public static Module1 Current => _this;

        /// <summary>
        /// Retrieve the singleton instance of the Log service.
        /// Guaranteed to be available after the module has been initialized.
        /// </summary>
        public static BIS_Log Log => _log;

        /// <summary>
        /// Retrieve the singleton instance of the business rules engine.
        /// Guaranteed to be available after the module has been initialized.
        /// </summary>
        public static IC_Rules IcRules => _icRules;

        /// <summary>
        /// Retrieve the singleton instance of the PostgreSQL database tools.
        /// Guaranteed to be available after the module has been initialized.
        /// </summary>
        public static BIS_DB_PostGre PostGreTool => _postGreTool;

        public static BisDbNjems NjemsTool => _njemsTool;
        public static BisDbCompass CompassTool => _compassTool;
        public static BisDbAccess AccessTool => _accessTool;

        /// <summary>
        /// Retrieve the singleton instance of the Regex service.
        /// </summary>
        public static Bis_Regex RegexTool => _regexTool;

        public static BisFileTools FileTool => _fileTool;

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
            // First, set the singleton instance of the module.
            _this = this;

            // Second, create instances of all core, shared services.
            // By doing this here, we guarantee they are ready before any UI
            // component (like a dockpane) is created or shown.

            // Initialize the Logger
            try
            {
            _log = new BIS_Log("IC_Loader_Pro");
            }
            catch(Exception ex){
                // This is a critical failure. The logger could not be created.
                // We can't use our logger, so we fall back to Debug output and a message box.
                System.Diagnostics.Debug.WriteLine($"FATAL: The Logging service failed to initialize. No further logging is possible. Exception: {ex}");
                ArcGIS.Desktop.Framework.Dialogs.MessageBox.Show($"The application could not start because the logging service failed to initialize.{Environment.NewLine}{ex.Message}", "Critical Initialization Error");
                return false; // Stop the initialization process.
            }

            System.Diagnostics.Debug.WriteLine("Module.Initialize: BIS_Log service created.");

            try
            {
                _fileTool = new BisFileTools(_log);
            }
            catch (Exception ex)
            {
                _log.RecordError("FATAL: Failed to create BisFileTools service.", ex, nameof(Initialize));
                return false;
            }

            // Initialize the Regex Tool
            try
            {
                _regexTool = new Bis_Regex(_log);
            }
            catch (Exception ex)
            {
                _log.RecordError("FATAL: Failed to create Bis_Regex service in Module.Initialize", ex, nameof(Initialize));
                return false;
            }

            // Initialize the Database Tool
            try
            {
                _postGreTool = new BIS_DB_PostGre(_log);
                System.Diagnostics.Debug.WriteLine("Module.Initialize: BIS_DB_PostGre service created successfully.");
            }
            catch (Exception ex)
            {
                // Write the full exception to both the debug output and the main log file.
                string errorMessage = $"Failed to create BIS_DB_PostGre service in Module.Initialize";
                System.Diagnostics.Debug.WriteLine($"FATAL: {errorMessage}");
                _log.RecordError($"FATAL: {errorMessage}",ex, nameof(Initialize));
            }

            _njemsTool = new BisDbNjems(_log);
            _compassTool = new BisDbCompass(_log);
            _accessTool = new BisDbAccess(_log);

            // Initialize the Rules Engine
            try
            {
                _icRules = new IC_Rules(_log, _postGreTool,_compassTool,_njemsTool,_accessTool , _fileTool,_regexTool);
                System.Diagnostics.Debug.WriteLine("Module.Initialize: IC_Rules service created successfully.");
            }
            catch (Exception ex)
            {
                // Write the full exception to both the debug output and the main log file.
                string errorMessage = $"Failed to create IC_Rules service in Module.Initialize";
                System.Diagnostics.Debug.WriteLine($"FATAL: {errorMessage}");
                _log.RecordError($"FATAL: {errorMessage}", ex, nameof(Initialize));
            }

            ProjectClosingEvent.Subscribe(OnProjectClosing);

            // Return a completed task.
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

        #endregion Overrides

    }
}
