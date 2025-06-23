using ArcGIS.Desktop.Framework.Contracts;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IC_Loader_Pro.Models
{
    // Ensure your class inherits from PropertyChangedBase from the Esri framework
    public class ICQueueSummary  : PropertyChangedBase
    {
        private string _name;
        public string Name
        {
            get => _name;
            // The SetProperty method handles the UI notification automatically
            set => SetProperty(ref _name, value);
        }

        private int _emailCount;
        public int EmailCount
        {
            get => _emailCount;
            set => SetProperty(ref _emailCount, value);
        }

        private int _passedCount;
        public int PassedCount
        {
            get => _passedCount;
            set => SetProperty(ref _passedCount, value);
        }

        private int _skippedCount;
        public int SkippedCount
        {
            get => _skippedCount;
            set => SetProperty(ref _skippedCount, value);
        }

        private int _failedCount;
        public int FailedCount
        {
            get => _failedCount;
            set => SetProperty(ref _failedCount, value);
        }
    }
}
