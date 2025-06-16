using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IC_Loader_Pro.Models
{
    internal class ActiveEmail
    {
        public string Subject { get; set; }
        public string From { get; set; }
        public List<string> ShapefilePaths { get; set; } // Paths to extracted shapefiles
        public object OriginalEmail { get; set; } // Optional: keep reference if needed
        public string PrefID { get; internal set; }
        public string DelID { get; internal set; }
    }
}
