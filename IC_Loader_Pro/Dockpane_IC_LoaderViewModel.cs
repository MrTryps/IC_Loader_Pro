using ArcGIS.Desktop.Framework;
using ArcGIS.Desktop.Framework.Contracts;
using ArcGIS.Desktop.Framework.Dialogs;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows.Input;


namespace IC_Loader_Pro
{
    // This class is now both the DockPane and the ViewModel
    internal class Dockpane_IC_LoaderViewModel : DockPane
    {
        private const string _dockPaneID = "IC_Loader_Pro_Dockpane_IC_Loader";

        #region Properties
        public ObservableCollection<Models.ICQueueInfo> ICQueues { get; } = new ObservableCollection<Models.ICQueueInfo>();

        private Models.ICQueueInfo _selectedQueue;
        public Models.ICQueueInfo SelectedQueue
        {
            get => _selectedQueue;
            set => SetProperty(ref _selectedQueue, value, () => SelectedQueue);
        }

        private Models.ActiveEmail _currentEmail;
        public Models.ActiveEmail CurrentEmail
        {
            get => _currentEmail;
            set => SetProperty(ref _currentEmail, value, () => CurrentEmail);
        }

        // Aggregate Counts
        private int _passedCount;
        public int PassedCount { get => _passedCount; set => SetProperty(ref _passedCount, value, () => PassedCount); }
        private int _skippedCount;
        public int SkippedCount { get => _skippedCount; set => SetProperty(ref _skippedCount, value, () => SkippedCount); }
        private int _failedCount;
        public int FailedCount { get => _failedCount; set => SetProperty(ref _failedCount, value, () => FailedCount); }
        #endregion

        #region Commands
        public ICommand ShowNotesCommand { get; }
        public ICommand SearchCommand { get; }
        public ICommand ToolsCommand { get; }
        public ICommand OptionsCommand { get; }
        #endregion

        /// <summary>
        /// The constructor is now responsible for creating the View
        /// and setting it as the dock pane's content.
        /// </summary>
        public Dockpane_IC_LoaderViewModel()
        {
            // Create an instance of the UserControl and set it as the Content.
            // The UserControl will automatically inherit this class instance as its DataContext.
            this.Content = new Dockpane_IC_LoaderView();

            // Initialize commands
            ShowNotesCommand = new RelayCommand(() => MessageBox.Show("Show Notes command executed."));
            SearchCommand = new RelayCommand(() => MessageBox.Show("Search command executed."));
            ToolsCommand = new RelayCommand(() => MessageBox.Show("Tools command executed."));
            OptionsCommand = new RelayCommand(() => MessageBox.Show("Options command executed."));

            LoadDummyData();        
        }


        private void LoadDummyData()
        {
            // It's good practice to clear the list first to prevent duplicates
            // if this method were ever called more than once.
            ICQueues.Clear();

            // Create and add some sample ICQueueInfo objects to the collection.
            // The UI will create one button for each item added here.
            ICQueues.Add(new Models.ICQueueInfo { Name = "CEAs", EmailCount = 12 });
            ICQueues.Add(new Models.ICQueueInfo { Name = "DNAs", EmailCount = 5 });
            ICQueues.Add(new Models.ICQueueInfo { Name = "WRAs", EmailCount = 21 });            

            // Set a default selection so the first button is active when the pane opens.
            SelectedQueue = ICQueues.FirstOrDefault();

            // You can also populate the other dummy data here if it's not already
            PassedCount = 11;
            SkippedCount = 2;
            FailedCount = 10;

            CurrentEmail = new Models.ActiveEmail
            {
                Subject = "FW: Institutional Control Submission - Site 123",
                PrefID = "g0000355",
                DelID = "gIS_1234"
            };
        }

        /// <summary>
        /// Show the DockPane.
        /// </summary>
        internal static void Show()
        {
            DockPane pane = FrameworkApplication.DockPaneManager.Find(_dockPaneID);
            if (pane == null)
                return;
            pane.Activate();
        }
    }
}