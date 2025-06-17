using ArcGIS.Desktop.Framework;
using ArcGIS.Desktop.Framework.Contracts;
using System.Linq;
using static BIS_Tools_2025_Core.BIS_Log;
using static IC_Loader_Pro.Module1;

namespace IC_Loader_Pro
{
    /// <summary>
    /// Button implementation to show the DockPane.
    /// </summary>
    internal class Dockpane_IC_Loader_ShowButton : Button
    {
        protected override void OnClick()
        {
            string dockpaneId = "IC_Loader_Pro_Dockpane_IC_Loader";
            Pane pane = FrameworkApplication.Panes.Find(dockpaneId)?.FirstOrDefault();
            if (pane == null)
            {
                // This should not happen, as the framework creates the pane based on the DAML.
                // We will log an error if the pane can't be found.
                Log.recordError($"Could not find dockpane with ID '{dockpaneId}'. Check Config.daml.", null, "ShowButton.OnClick");
                return;
            }

            // Activate the dockpane to make it visible and bring it to the front.
            pane.Activate();
        }
    }
}
