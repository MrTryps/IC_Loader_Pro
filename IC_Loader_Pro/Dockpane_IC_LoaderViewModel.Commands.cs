using ArcGIS.Desktop.Framework.Contracts;
using ArcGIS.Desktop.Framework.Threading.Tasks;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Input;
using static BIS_Tools_2025_Core.BIS_Log;
using static IC_Loader_Pro.Module1;

namespace IC_Loader_Pro
{
    internal partial class Dockpane_IC_LoaderViewModel : DockPane
    {
        #region Commands
        public ICommand SaveCommand { get; private set; }
        public ICommand SkipCommand { get; private set; }
        public ICommand RejectCommand { get; private set; }
        public ICommand ShowNotesCommand { get; private set; }
        public ICommand SearchCommand { get; private set; }
        public ICommand ToolsCommand { get; private set; }
        public ICommand OptionsCommand { get; private set; }
        public ICommand RefreshQueuesCommand { get; private set; }
        #endregion

        #region Command Methods
        private async Task OnSave()
        {
            Log.recordMessage("Save button was clicked.", Bis_Log_Message_Type.Note);

            // We can now use 'await' here for any async GIS work in the future,
            // for example, saving features to the 'manually_added' layer.
            await QueuedTask.Run(() => {
                // ... future GIS logic ...
            });
        }

        // Also update OnSkip and OnReject
        private async Task OnSkip()
        {
            Log.recordMessage("Skip button was clicked.", Bis_Log_Message_Type.Note);
            // Use Task.CompletedTask as a placeholder since there's no async work yet.
            await Task.CompletedTask;
        }

        private async Task OnReject()
        {
            Log.recordMessage("Reject button was clicked.", Bis_Log_Message_Type.Note);
            await Task.CompletedTask;
        }

        private async Task OnShowNotes()
        {
            Log.recordMessage("Menu: Notes was clicked.", Bis_Log_Message_Type.Note);
            Log.open();
            await Task.CompletedTask;
        }

        private async Task OnSearch()
        {
            Log.recordMessage("Menu: Search was clicked.", Bis_Log_Message_Type.Note);
            await Task.CompletedTask;
        }

        private async Task OnTools()
        {
            Log.recordMessage("Menu: Tools was clicked.", Bis_Log_Message_Type.Note);
            await Task.CompletedTask;
        }

        private async Task OnOptions()
        {
            Log.recordMessage("Menu: Options was clicked.", Bis_Log_Message_Type.Note);
            await Task.CompletedTask;
        }

        /// <summary>
        /// This method will contain the logic to call your Outlook library and get the real data.
        /// </summary>
        private Task RefreshICQueuesAsync()
        {
            return QueuedTask.Run(() =>
            {
                // This lock ensures that the collection is not modified by two threads at once.
                lock (_lockQueueCollection)
                {
                    _listOfQueues.Clear();

                    // For now, we use sample data. Later, we will replace this with a call
                    // to your BIS_IC_InputClasses_2025 library.
                    _listOfQueues.Add(new Models.ICQueueInfo { Name = "CEAs", EmailCount = 12, PassedCount = 0, SkippedCount = 0, FailedCount = 0 });
                    _listOfQueues.Add(new Models.ICQueueInfo { Name = "DNAs", EmailCount = 5, PassedCount = 0, SkippedCount = 0, FailedCount = 0 });
                    _listOfQueues.Add(new Models.ICQueueInfo { Name = "WRAs", EmailCount = 21, PassedCount = 0, SkippedCount = 0, FailedCount = 0 });
                }

                // Select the first item by default
                SelectedQueue = _readOnlyListOfQueues.FirstOrDefault();
            });
        }
        #endregion
    }
}
