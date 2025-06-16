using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IC_Loader_Pro
{
    internal class TestData : ArcGIS.Desktop.Framework.Contracts.PropertyChangedBase
    {
        private string _id;
        public string Id
        {
            get { return _id; }
            set
            { SetProperty(ref _id, value, () => Id);
            }
        }

        private string _label;
        public string Label
        {
            get { return _label; }
            set
            {
                SetProperty(ref _label, value, () => _label);
            }
        }

        private int _num;
        public int Num
        {
            get { return _num; }
            set
            {
                SetProperty(ref _num, value, () => _num);
            }
        }

        internal TestData(string id, string label, int num)
        {
            Id = id;
            Label = label;
            Num = num;
        }

       
    }
}
