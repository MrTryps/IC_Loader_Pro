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
            Cursor = Cursors.Arrow;
        }

        /// <summary>
        /// Called by the framework when the tool is activated.
        /// </summary>
        protected override Task OnToolActivateAsync(bool active)
        {
            if (active)
            {
                Log.RecordMessage("SelectShapeTool has been ACTIVATED.", BIS_Log.BisLogMessageType.Note);
            }
            else
            {
                Log.RecordMessage("SelectShapeTool has been DEACTIVATED.", BIS_Log.BisLogMessageType.Note);
            }
            return base.OnToolActivateAsync(active);
        }

        /// <summary>
        /// Overrides the mouse down event to perform a spatial query.
        /// </summary>
        protected override Task HandleMouseDownAsync(MapViewMouseButtonEventArgs e)
        {
            Log.RecordMessage("HandleMouseDownAsync triggered.", BIS_Log.BisLogMessageType.Note);
            // We only care about the left mouse button.
            if (e.ChangedButton != MouseButton.Left)
            {
                return Task.CompletedTask;
            }
            Log.RecordMessage("Left button triggered.", BIS_Log.BisLogMessageType.Note);
            var pane = FrameworkApplication.DockPaneManager.Find(Dockpane_IC_LoaderViewModel.DockPaneId) as Dockpane_IC_LoaderViewModel;
            if (pane == null)
                return Task.CompletedTask;

            return QueuedTask.Run(() =>
            {
                
                var graphicsLayer = pane.GetGraphicsLayer();
                if (graphicsLayer == null)
                {
                    Log.RecordMessage("Could not get graphics layer from ViewModel.", BIS_Log.BisLogMessageType.Note);
                    return;

                }

                // 1. Convert the screen point to a map point.
                MapPoint mapPoint = MapView.Active.ClientToMap(e.ClientPoint);
                if (mapPoint == null)
                {
                    Log.RecordMessage("Could not convert click point to map point.", BIS_Log.BisLogMessageType.Note);
                    return;
                }

                // 2. Define a small search tolerance in map units.
                //    This makes it easier for the user to click a shape.
                double searchTolerance = 5.0; //MapView.Active.Extent.Width / 1000;
                Polygon searchBuffer = GeometryEngine.Instance.Buffer(mapPoint, searchTolerance) as Polygon;

                // 3. Find the first element whose geometry intersects our search buffer.
                var topElement = graphicsLayer.GetElements()
             .OfType<GraphicElement>()
             .FirstOrDefault(graphicElement =>
             {
                 // Get the base CIMGraphic
                 var cimGraphic = graphicElement.GetGraphic();
                 // Cast it to the specific type for polygons
                 var polygonGraphic = cimGraphic as CIMPolygonGraphic;
                 if (polygonGraphic == null)
                     return false; // It's not a polygon, so we can't test it.

                 // Now, perform the intersect test on the .Polygon property
                 return GeometryEngine.Instance.Intersects(polygonGraphic.Polygon, searchBuffer);
             });

                if (topElement != null)
                {
                    Log.RecordMessage($"SUCCESS: Found intersecting element with Name: '{topElement.Name}'", BIS_Log.BisLogMessageType.Note);
                    pane.SelectShapeFromTool(topElement.Name);
                    pane.DeactivateSelectTool();
                }
                else
                {
                    Log.RecordMessage("INFO: Spatial query ran, but no intersecting element was found.", BIS_Log.BisLogMessageType.Note);
                    pane.DeactivateSelectTool();
                }
            });
        }
    }
}