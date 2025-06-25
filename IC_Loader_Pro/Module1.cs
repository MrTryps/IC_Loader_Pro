using ArcGIS.Desktop.Framework;
using ArcGIS.Desktop.Framework.Contracts;
using BIS_Tools_2025_Core;
using IC_Rules_2025;
using System;
using System.Threading.Tasks;


namespace IC_Loader_Pro
{
    internal class Module1 : Module
    {
        private static Module1 _this = null;
        private static BIS_Log _log;
        private static IC_Rules _icRules = null;
        private static BIS_DB_Tools.BIS_DB_PostGre _postGreTool = null;

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
        public static BIS_DB_Tools.BIS_DB_PostGre PostGreTool => _postGreTool;

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

  // Initialize the Database Tool
            try
            {
                _postGreTool = new BIS_DB_Tools.BIS_DB_PostGre(_log);
                System.Diagnostics.Debug.WriteLine("Module.Initialize: BIS_DB_PostGre service created successfully.");
            }
            catch (Exception ex)
            {
                // Write the full exception to both the debug output and the main log file.
                string errorMessage = $"Failed to create BIS_DB_PostGre service in Module.Initialize";
                System.Diagnostics.Debug.WriteLine($"FATAL: {errorMessage}");
                _log.recordError($"FATAL: {errorMessage}",ex, nameof(Initialize));
            }
            // Initialize the Rules Engine
            try
            {
                _icRules = new IC_Rules(_log, _postGreTool);
                System.Diagnostics.Debug.WriteLine("Module.Initialize: IC_Rules service created successfully.");
            }
            catch (Exception ex)
            {
                // Write the full exception to both the debug output and the main log file.
                string errorMessage = $"Failed to create IC_Rules service in Module.Initialize";
                System.Diagnostics.Debug.WriteLine($"FATAL: {errorMessage}");
                _log.recordError($"FATAL: {errorMessage}", ex, nameof(Initialize));
            }
          

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

        #endregion Overrides

    }
}
