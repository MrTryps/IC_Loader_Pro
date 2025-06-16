using ArcGIS.Desktop.Framework.Contracts;


namespace IC_Loader_Pro
{
    /// <summary>
    /// Button implementation to show the DockPane.
    /// </summary>
    internal class Dockpane_IC_Loader_ShowButton : Button
    {
        protected override void OnClick()
        {
            Dockpane_IC_LoaderViewModel.Show();
        }
    }
}
