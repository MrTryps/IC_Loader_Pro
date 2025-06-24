using ArcGIS.Core.Geometry; // Required for Polygon
using ArcGIS.Desktop.Framework.Contracts;

namespace IC_Loader_Pro.Models
{
    public class ShapeItem : PropertyChangedBase
    {
        private Polygon _geometry;
        /// <summary>
        /// The actual polygon geometry of the shape.
        /// </summary>
        public Polygon Geometry
        {
            get => _geometry;
            set => SetProperty(ref _geometry, value);
        }

        private string _sourceFile;
        /// <summary>
        /// The name of the original file this shape came from (e.g., "attachment1.shp").
        /// </summary>
        public string SourceFile
        {
            get => _sourceFile;
            set => SetProperty(ref _sourceFile, value);
        }

        private bool _isValid;
        /// <summary>
        /// A flag indicating if the shape passed our validation rules (is a polygon, has area, etc.).
        /// </summary>
        public bool IsValid
        {
            get => _isValid;
            set => SetProperty(ref _isValid, value);
        }

        private string _validationMessage = "OK";
        /// <summary>
        /// A message explaining why a shape is not valid.
        /// </summary>
        public string ValidationMessage
        {
            get => _validationMessage;
            set => SetProperty(ref _validationMessage, value);
        }
    }
}