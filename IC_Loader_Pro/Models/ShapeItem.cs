using ArcGIS.Core.Geometry; // Required for Polygon
using ArcGIS.Desktop.Framework.Contracts;
using System.Collections.Generic;

namespace IC_Loader_Pro.Models
{
    public class ShapeItem : PropertyChangedBase
    {
        // --- Properties for UI state ---
        private bool _isShownInMap = true;
        public bool IsShownInMap
        {
            get => _isShownInMap;
            set => SetProperty(ref _isShownInMap, value);
        }

        private bool _isSelectedForUse = false;
        public bool IsSelectedForUse
        {
            get => _isSelectedForUse;
            set => SetProperty(ref _isSelectedForUse, value);
        }

        // --- Core Shape Properties ---
        public int ShapeReferenceId { get; set; } // The original OBJECTID
        public Polygon Geometry { get; set; }
        public string SourceFile { get; set; }
        public string ShapeType { get; set; } // e.g., "Polygon", "Polyline"
        public double Area { get; set; }

        // --- Validation Properties ---
        public bool IsValid { get; set; }
        public string Status { get; set; } // e.g., "OK", "Self-Intersecting"

        /// <summary>
        /// A dictionary to store the attribute values for the fields defined
        /// in the IC_Rules (the "fields to mine").
        /// </summary>
        public Dictionary<string, object> Attributes { get; set; } = new Dictionary<string, object>();
    }
}