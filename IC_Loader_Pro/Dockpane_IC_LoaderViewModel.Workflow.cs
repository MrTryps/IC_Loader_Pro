using ArcGIS.Desktop.Framework.Threading.Tasks;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using static IC_Loader_Pro.Module1;
using IC_Loader_Pro.Services;
using IC_Loader_Pro.Models;
using BIS_Tools_DataModels_2025;
using static BIS_Tools_2025_Core.BIS_Log;

namespace IC_Loader_Pro
{
    internal partial class Dockpane_IC_LoaderViewModel
    {
        /// <summary>
        /// Asynchronously fetches summary data for each IC Queue from the OutlookService
        /// and populates the UI's collection of toggle buttons.
        /// </summary>
        private async Task RefreshICQueuesAsync()
        {
            // Log the start of the operation and update the UI status.
            Log.recordMessage("Refreshing IC Queue summaries from source...", Bis_Log_Message_Type.Note);
            StatusMessage = "Loading email queues...";
            var rulesEngine = Module1.IcRules;
            try
            {
                // This work will be done on a background thread to keep the UI responsive.
                var summaryList = await QueuedTask.Run( () =>
                {
                    var outlookService = new OutlookService();
                    var summaries = new List<ICQueueSummary>();

                    // Get the list of IC Types to process from your rules engine.                 
                    foreach (string icType in IcRules.ReturnIcTypes())
                    {
                        try
                        {
                            // Get the specific settings for this queue, including the folder name.
                            IcGisTypeSetting icSetting = rulesEngine.ReturnIcGisTypeSettings(icType);
                            string outlookFolderPath = icSetting.OutlookInboxFolderPath;

                            // Call our service to get the detailed list of emails for this folder.
                            List<EmailItem> emailsInQueue = outlookService.GetEmailsFromFolderPath(outlookFolderPath);

                            // Create the summary object from the results.
                            summaries.Add(new ICQueueSummary
                            {
                                Name = icType,
                                EmailCount = emailsInQueue.Count,
                                PassedCount = 0, // Will be calculated later.
                                SkippedCount = 0,
                                FailedCount = 0
                            });
                        }
                        catch (Exception ex)
                        {
                            // Log the error for any specific queue that fails, then continue to the next.
                            Log.recordError($"An error occurred while processing queue '{icType}'.", ex, nameof(RefreshICQueuesAsync));
                        }
                    }
                    return summaries;
                });

                // Now that we have the data, update the UI's collection.
                // Because we enabled collection synchronization in the constructor,
                // we can safely modify our private list and the UI will update automatically.
                lock (_lockQueueCollection)
                {
                    _ListOfIcEmailTypeSummaries.Clear();
                    foreach (var summary in summaryList)
                    {
                        _ListOfIcEmailTypeSummaries.Add(summary);
                    }
                }

                // Set the default selected item.
                SelectedIcType = PublicListOfIcEmailTypeSummaries.FirstOrDefault();
                Log.recordMessage($"Successfully loaded {PublicListOfIcEmailTypeSummaries.Count} queues.", Bis_Log_Message_Type.Note);
            }
            catch (Exception ex)
            {
                // This will catch any unexpected errors in the overall process.
                Log.recordError("A fatal error occurred while refreshing the IC Queues.", ex, nameof(RefreshICQueuesAsync));
                StatusMessage = "Error loading email queues.";
            }
        }
    }
}
