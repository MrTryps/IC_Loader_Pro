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
using System.Windows.Input;
using BIS_Tools_2025_Core;
using System.Data;
using IC_Rules_2025;

namespace IC_Loader_Pro
{
    internal class Module1 : Module
    {
        private static Module1 _this = null;
        private static BIS_Log Log;
        private static IC_Rules IC_Rules = null;
        private static BIS_DB_Tools.BIS_DB_PostGre PostGreTool = null;

        /// <summary>
        /// Retrieve the singleton instance to this module here
        /// </summary>
        /// 

        public static Module1 Current => _this ??= (Module1)FrameworkApplication.FindModule("IC_Loader_Pro_Module");
        public static BIS_Log _Log => Log ??= new BIS_Log("IC_Loader_Pro");
        public static IC_Rules _IC_Rules => IC_Rules ??= new IC_Rules(Log);
        public static BIS_DB_Tools.BIS_DB_PostGre _PostGreTool => PostGreTool ??= new BIS_DB_Tools.BIS_DB_PostGre(Log);

        #region Overrides
        /// <summary>
        /// Called by Framework when ArcGIS Pro is closing
        /// </summary>
        /// <returns>False to prevent Pro from closing, otherwise True</returns>
        protected override bool CanUnload()
        {
            //TODO - add your business logic
            //return false to ~cancel~ Application close
            return true;
        }

        #endregion Overrides

    }
}
