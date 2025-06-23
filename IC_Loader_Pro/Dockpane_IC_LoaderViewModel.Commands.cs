using ArcGIS.Desktop.Framework;
using ArcGIS.Desktop.Framework.Contracts;
using ArcGIS.Desktop.Framework.Threading.Tasks;
using BIS_Tools_DataModels_2025;
using IC_Loader_Pro.Models;
using IC_Loader_Pro.Services;
using Microsoft.Office.Interop.Outlook;
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
                // Instantiate the service once, outside the loop, for efficiency.
                var outlookService = new OutlookService();

                // We'll build a temporary list here on the background thread.
                var summaryList = new List<ICQueueSummary>();

                foreach (string IcType in IC_Rules.ReturnIcTypes())
                {
                    try
                    {
                        IcGisTypeSetting icSetting = IC_Rules.ReturnIcGisTypeSettings(IcType);
                        string outlookFolderName = icSetting.EmailFolderSet.InboxFolderName;

                        // 1. Call our service to get the detailed list of emails for this queue.
                        List<EmailItem> emailsInQueue = outlookService.GetEmailsFromSubfolder(outlookFolderName);

                        // 2. Create the summary object using the results from the service call.
                        var summary = new ICQueueSummary
                        {
                            Name = IcType,
                            EmailCount = emailsInQueue.Count,
                            PassedCount = 0,  // This will be calculated later as the user works through the queue.
                            SkippedCount = 0, // This will be calculated later.
                            FailedCount = 0   // This will be calculated later.
                        };

                        summaryList.Add(summary);
                    }
                    catch (System.Exception ex)
                    {
                        // Log the error for the specific queue that failed, then continue to the next.
                        Log.recordError($"An error occurred while checking queue '{IcType}'.", ex, nameof(RefreshICQueuesAsync));
                    }
                }

                // 3. Now that we have all the data, update the main UI collection on the UI thread.
                // This is safer and more efficient than updating it inside the loop.
                FrameworkApplication.Current.Dispatcher.Invoke(() =>
                {
                    lock (_lockQueueCollection) // Use the lock for thread safety
                    {
                        _ListOfIcEmailTypeSummaries.Clear();
                        foreach (var summary in summaryList)
                        {
                            _ListOfIcEmailTypeSummaries.Add(summary);
                        }
                    }

                    // Select the first item in the list by default
                    SelectedQueue = _readOnly_ListOfIcEmailTypeSummaries.FirstOrDefault();
                });
            });
        }
        #endregion
    }
}
