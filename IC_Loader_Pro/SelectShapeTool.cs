using ArcGIS.Core.CIM;
using ArcGIS.Core.Geometry;
using ArcGIS.Desktop.Framework;
using ArcGIS.Desktop.Framework.Threading.Tasks;
using ArcGIS.Desktop.Layouts;
using ArcGIS.Desktop.Mapping;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Input;
using static IC_Loader_Pro.Module1;

namespace IC_Loader_Pro
{
    /// <summary>
    /// A custom MapTool to handle clicking on shapes in our graphics layer.
    /// This is the correct architectural pattern for this functionality.
    /// </summary>
    internal class SelectShapeTool : MapTool
    {
        public SelectShapeTool()
        {
            // Set a standard arrow cursor for the tool.
            Cursor = Cursors.Cross;
        }

        /// <summary>
        /// Called by the framework when the tool is activated.
        /// </summary>
        protected override Task OnToolActivateAsync(bool hasMapViewChanged)
        {
            Cursor = Cursors.Cross;
            Log.RecordMessage("SelectShapeTool has been ACTIVATED.", BIS_Log.BisLogMessageType.Note);
            return base.OnToolActivateAsync(hasMapViewChanged);
        }

        protected override Task OnToolDeactivateAsync(bool hasMapViewChanged)
        {
            Log.RecordMessage("SelectShapeTool has been DEACTIVATED.", BIS_Log.BisLogMessageType.Note);
            return base.OnToolDeactivateAsync(hasMapViewChanged);
        }

        /// <summary>
        /// This synchronous method is called first for any mouse down event.
        /// We set args.Handled = true to indicate that we are taking control
        /// and that the framework should now call HandleMouseDownAsync.
        /// </summary>
        protected override void OnToolMouseDown(MapViewMouseButtonEventArgs args)
        {
            // We only handle the left mouse button.
            if (args.ChangedButton == MouseButton.Left)
            {
                args.Handled = true;
            }
        }

        protected override Task HandleMouseDownAsync(MapViewMouseButtonEventArgs e)
        {
            if (e.ChangedButton != MouseButton.Left) // if (e.ChangedButton != MouseButton.Left || e.Action != MouseButton.Down
                return Task.CompletedTask;

            var pane = FrameworkApplication.DockPaneManager.Find(Dockpane_IC_LoaderViewModel.DockPaneId) as Dockpane_IC_LoaderViewModel;
            if (pane == null)
                return Task.CompletedTask;

            // Check if the Control key is held down BEFORE the QueuedTask.
            bool isCtrlKeyDown = (Keyboard.IsKeyDown(Key.LeftCtrl) || Keyboard.IsKeyDown(Key.RightCtrl));

            return QueuedTask.Run(() =>
            {
                Log.RecordMessage("HandleMouseDownAsync triggered in SelectShapeTool.", BIS_Log.BisLogMessageType.Note);

                var graphicsLayer = pane.GetGraphicsLayer();
                if (graphicsLayer == null) return;

                MapPoint mapPoint = MapView.Active.ClientToMap(e.ClientPoint);
                if (mapPoint == null) return;

                double searchTolerance = MapView.Active.Extent.Width / 1000;
                Polygon searchBuffer = GeometryEngine.Instance.Buffer(mapPoint, searchTolerance) as Polygon;

                var topElement = graphicsLayer.GetElements()
                    .OfType<GraphicElement>()
                    .FirstOrDefault(graphicElement =>
                    {
                        var polygonGraphic = graphicElement.GetGraphic() as CIMPolygonGraphic;
                        if (polygonGraphic == null) return false;
                        return GeometryEngine.Instance.Intersects(polygonGraphic.Polygon, searchBuffer);
                    });

                if (topElement != null)
                {
                    if (isCtrlKeyDown)
                    {
                        // If Ctrl is down, toggle the selection and keep the tool active.
                        Log.RecordMessage($"SUCCESS: Toggling selection for element '{topElement.Name}'.", BIS_Log.BisLogMessageType.Note);
                        pane.ToggleShapeSelectionFromTool(topElement.Name);
                    }
                    else
                    {
                        // If Ctrl is NOT down, replace the selection.
                        Log.RecordMessage($"SUCCESS: Setting selection to element '{topElement.Name}'.", BIS_Log.BisLogMessageType.Note);
                        pane.SelectShapeFromTool(topElement.Name);

                        // ** ADDED LOGIC: Deactivate the tool after a single selection. **
                        Log.RecordMessage("Single selection complete. Deactivating tool.", BIS_Log.BisLogMessageType.Note);
                        pane.DeactivateSelectTool();
                    }
                }
                else
                {
                    Log.RecordMessage("INFO: No intersecting element was found at the clicked point.", BIS_Log.BisLogMessageType.Note);
                }
            });
        }

        /// <summary>
        /// This synchronous method is called first for any key up event.
        /// We set args.Handled = true to indicate that we want to take control
        /// and that the framework should now call HandleKeyUpAsync.
        /// </summary>
        protected override void OnToolKeyUp(MapViewKeyEventArgs args)
        {
            // We only care about the Ctrl keys.
            if (args.Key == Key.LeftCtrl || args.Key == Key.RightCtrl)
            {
                args.Handled = true;
            }
        }

        /// <summary>
        /// Overrides the key up event. If the user releases the Control key,
        /// we will deactivate this tool and revert to the default explore tool.
        /// </summary>
        protected override Task HandleKeyUpAsync(MapViewKeyEventArgs e)
        {         
            Log.RecordMessage("Ctrl key released, deactivating SelectShapeTool.", BIS_Log.BisLogMessageType.Note);

            // Find our dockpane
            var pane = FrameworkApplication.DockPaneManager.Find(Dockpane_IC_LoaderViewModel.DockPaneId) as Dockpane_IC_LoaderViewModel;
            if (pane != null)
            {
                // Tell the ViewModel to deactivate the tool. This will un-check the
                // toggle button and cause the explore tool to become active.
                pane.DeactivateSelectTool();
            }
            
            return base.HandleKeyUpAsync(e);
        }
    }
}