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
    internal class ShowProWfrm_LoadICs : Button
    {

        private ProWfrm_LoadICs _prowfrm_loadics = null;

        protected override void OnClick()
        {
            //already open?
            if (_prowfrm_loadics != null)
                return;
            _prowfrm_loadics = new ProWfrm_LoadICs();
            _prowfrm_loadics.Owner = FrameworkApplication.Current.MainWindow;
            _prowfrm_loadics.Closed += (o, e) => { _prowfrm_loadics = null; };
            _prowfrm_loadics.Show();
            //uncomment for modal
            //_prowfrm_loadics.ShowDialog();
        }

    }
}
