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
    internal class Showfrm_Show_Ics_to_Load : Button
    {

        private frm_Show_Ics_to_Load _frm_show_ics_to_load = null;

        protected override void OnClick()
        {
            //already open?
            if (_frm_show_ics_to_load != null)
                return;
            _frm_show_ics_to_load = new frm_Show_Ics_to_Load();
            _frm_show_ics_to_load.Owner = FrameworkApplication.Current.MainWindow;
            _frm_show_ics_to_load.Closed += (o, e) => { _frm_show_ics_to_load = null; };
            _frm_show_ics_to_load.Show();
            //uncomment for modal
            //_frm_show_ics_to_load.ShowDialog();
        }

    }
}
