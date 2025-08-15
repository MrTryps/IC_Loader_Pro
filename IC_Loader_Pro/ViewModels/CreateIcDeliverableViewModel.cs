// In IC_Loader_Pro/ViewModels/CreateIcDeliverableViewModel.cs

using ArcGIS.Desktop.Catalog;
using ArcGIS.Desktop.Core;
using ArcGIS.Desktop.Framework;
using ArcGIS.Desktop.Framework.Contracts;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Input;

namespace IC_Loader_Pro.ViewModels
{
    public class CreateIcDeliverableViewModel : PropertyChangedBase
    {
        #region Private Members
        private string _selectedIcType;
        private string _prefId;
        private string _gisFilePath;
        private string _validationStatus;
        private bool _isPrefIdValid;
        #endregion

        #region Public Properties for Binding
        /// <summary>
        /// The list of available IC Types to populate the dropdown.
        /// </summary>
        public List<string> IcTypes { get; }

        public string SelectedIcType
        {
            get => _selectedIcType;
            set
            {
                SetProperty(ref _selectedIcType, value);
                (CreateDeliverableCommand as RelayCommand)?.RaiseCanExecuteChanged();
            }
        }

        public string PrefId
        {
            get => _prefId;
            set
            {
                SetProperty(ref _prefId, value);
                // When PrefId changes, reset the validation status
                IsPrefIdValid = false;
                ValidationStatus = "Not yet validated.";
                (ValidatePrefIdCommand as RelayCommand)?.RaiseCanExecuteChanged();
                (CreateDeliverableCommand as RelayCommand)?.RaiseCanExecuteChanged();
            }
        }

        public string GisFilePath
        {
            get => _gisFilePath;
            set
            {
                SetProperty(ref _gisFilePath, value);
                (CreateDeliverableCommand as RelayCommand)?.RaiseCanExecuteChanged();
            }
        }

        public string ValidationStatus
        {
            get => _validationStatus;
            set => SetProperty(ref _validationStatus, value);
        }

        public bool IsPrefIdValid
        {
            get => _isPrefIdValid;
            set => SetProperty(ref _isPrefIdValid, value);
        }

        public ICommand ValidatePrefIdCommand { get; }
        public ICommand BrowseForFileCommand { get; }
        // Renamed for consistency
        public ICommand CreateDeliverableCommand { get; }
        #endregion

        public CreateIcDeliverableViewModel()
        {
            // Populate the list of IC Types from the rules engine
            IcTypes = Module1.IcRules.ReturnIcTypes();
            SelectedIcType = IcTypes.FirstOrDefault();

            ValidationStatus = "Not yet validated.";

            // Initialize commands
            ValidatePrefIdCommand = new RelayCommand(async () => await OnValidatePrefIdAsync(), () => !string.IsNullOrWhiteSpace(PrefId) && !IsPrefIdValid);
            BrowseForFileCommand = new RelayCommand(async () => await OnBrowseForFileAsync());
            CreateDeliverableCommand = new RelayCommand(() => { /* This command is handled by the View */ },
                () => IsPrefIdValid && !string.IsNullOrEmpty(GisFilePath) && !string.IsNullOrEmpty(SelectedIcType));
        }

        private async Task OnValidatePrefIdAsync()
        {
            ValidationStatus = "Validating...";
            bool isValid = false;

            // Run the database check on a background thread
            await Task.Run(() =>
            {
                isValid = Module1.IcRules.IsValidPrefId(PrefId);
            });

            IsPrefIdValid = isValid;
            ValidationStatus = IsPrefIdValid ? "Pref ID is valid." : "Pref ID not found.";

            // Notify the UI that the command states may have changed
            (ValidatePrefIdCommand as RelayCommand)?.RaiseCanExecuteChanged();
            (CreateDeliverableCommand as RelayCommand)?.RaiseCanExecuteChanged();
        }

        private async Task OnBrowseForFileAsync()
        {
            //var browseFilter = new BrowseProjectFilter("esri_browseDialogFilters_shapefiles_all", "esri_browseDialogFilters_cad_all")
            //{
            //    Name = "GIS Files (Shapefile, DWG)"
            //};

            var openDialog = new OpenItemDialog
            {
                Title = "Select Submission Fileset",
                MultiSelect = false,
               // BrowseFilter = browseFilter
            };

            if (openDialog.ShowDialog() == true)
            {
                GisFilePath = openDialog.Items.FirstOrDefault()?.Path;
            }
        }
    }
}