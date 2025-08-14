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

        private bool _isHidden = false;
        public bool IsHidden
        {
            get => _isHidden;
            set => SetProperty(ref _isHidden, value);
        }

        private string _description = string.Empty;
        public string Description
        {
            get => _description;
            set => SetProperty(ref _description, value);
        }

        /// <summary>
        /// A flag indicating that this shape's layer name matched a filter
        /// and should be automatically moved to the 'use' list.
        /// </summary>
        public bool IsAutoSelected { get; set; } = false;

        // --- Core Shape Properties ---
        public int ShapeReferenceId { get; set; } // The original OBJECTID
        public Polygon Geometry { get; set; }
        public string SourceFile { get; set; }
        public string ShapeType { get; set; } // e.g., "Polygon", "Polyline"
        public double Area { get; set; }

        // --- Validation Properties ---
        public bool IsValid { get; set; } = true;
        public string Status { get; set; } // e.g., "OK", "Self-Intersecting"

        private double _distanceFromSite;
        /// <summary>
        /// The calculated distance from the site's coordinates to this shape.
        /// </summary>
        public double DistanceFromSite
        {
            get => _distanceFromSite;
            set => SetProperty(ref _distanceFromSite, value);
        }


        /// <summary>
        /// A dictionary to store the attribute values for the fields defined
        /// in the IC_Rules (the "fields to mine").
        /// </summary>
        public Dictionary<string, object> Attributes { get; set; } = new Dictionary<string, object>();
    }
}