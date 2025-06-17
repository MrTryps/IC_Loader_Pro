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
        public const string DockPaneId = "IC_Loader_Pro_Dockpane_IC_Loader";
        protected override void OnClick()
        {
            // 1. Call the static method to ensure the pane is visible and active.
            Dockpane_IC_LoaderViewModel.Show();

            // 2. Find the ViewModel instance so we can call our one-time initialization method.
            var vm = FrameworkApplication.DockPaneManager.Find(Dockpane_IC_LoaderViewModel.DockPaneId) as Dockpane_IC_LoaderViewModel;
            if (vm != null)
            {
                // This triggers the logic to check/create the map and layers.
                _ = vm.LoadAndInitializeAsync();
            }
        }
    }
}
