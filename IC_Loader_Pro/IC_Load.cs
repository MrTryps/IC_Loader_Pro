using ArcGIS.Core.CIM;
using ArcGIS.Core.Data;
using ArcGIS.Core.Geometry;
using ArcGIS.Desktop.Catalog;
using ArcGIS.Desktop.Core;
using ArcGIS.Desktop.Editing;
using ArcGIS.Desktop.Extensions;
using ArcGIS.Desktop.Framework;
using ArcGIS.Desktop.Framework.Contracts;
using ArcGIS.Desktop.Framework.Dialogs;
using ArcGIS.Desktop.Framework.Threading.Tasks;
using ArcGIS.Desktop.Layouts;
using ArcGIS.Desktop.Mapping;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IC_Loader_Pro
{
    internal class IC_Load : Button
    {
        protected override void OnClick()
        {   
            var existingWindow = FrameworkApplication.Current.Windows
                       .OfType<frm_Show_Ics_to_Load>()
                       .FirstOrDefault();

            if (existingWindow != null)
            {
                // Window already exists, just bring it to the front
                existingWindow.Activate();
                return; // Stop here, don't create a new one
            }
        }
    }
}
