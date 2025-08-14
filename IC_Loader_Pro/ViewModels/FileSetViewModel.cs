using ArcGIS.Core.Internal.CIM;
using ArcGIS.Desktop.Framework.Contracts;
using BIS_Tools_DataModels_2025;
using System.Collections.Generic;
using System.Linq;

namespace IC_Loader_Pro.ViewModels
{
    public class FileSetViewModel : PropertyChangedBase
    {
        internal readonly fileset Model;

        public string FileName => Model.fileName;
        public string FileSetType => Model.filesetType;
        public bool IsValid => Model.validSet;
        public List<string> Extensions => Model.extensions;

        private int _totalFeatureCount;
        public int TotalFeatureCount
        {
            get => _totalFeatureCount;
            set => SetProperty(ref _totalFeatureCount, value);
        }

        private int _validFeatureCount;
        public int ValidFeatureCount
        {
            get => _validFeatureCount;
            set => SetProperty(ref _validFeatureCount, value);
        }

        private int _invalidFeatureCount;
        public int InvalidFeatureCount
        {
            get => _invalidFeatureCount;
            set => SetProperty(ref _invalidFeatureCount, value);
        }

        private bool _isLoadedInMap;
        public bool IsLoadedInMap
        {
            get => _isLoadedInMap;
            set => SetProperty(ref _isLoadedInMap, value);
        }
        /// <summary>
        /// A formatted string of the extensions for use in a tooltip.
        /// </summary>
        public string ExtensionsTooltip => $"Contains: {string.Join(", ", Extensions)}";

        private string _submissionId;
        /// <summary>
        /// The unique Submission ID (Sub_ID) assigned after this fileset
        /// is recorded in the database.
        /// </summary>
        public string SubmissionId
        {
            get => _submissionId;
            set => SetProperty(ref _submissionId, value);
        }


        public FileSetViewModel(fileset model)
        {
            Model = model;
        }
    }
}