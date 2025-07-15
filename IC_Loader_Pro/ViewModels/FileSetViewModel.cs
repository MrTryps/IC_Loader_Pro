using ArcGIS.Core.Internal.CIM;
using ArcGIS.Desktop.Framework.Contracts;
using BIS_Tools_DataModels_2025; // Reference your core data model

namespace IC_Loader_Pro.ViewModels
{
    public class FileSetViewModel : PropertyChangedBase
    {
        private readonly Fileset _model;

        public string FileName => _model.FileName;
        public string FileSetType => _model.FilesetType;
        public bool IsValid => _model.ValidSet;

        public FileSetViewModel(Fileset model)
        {
            _model = model;
        }
    }
}