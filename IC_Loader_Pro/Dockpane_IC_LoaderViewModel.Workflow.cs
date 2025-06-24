using ArcGIS.Desktop.Framework.Threading.Tasks;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static BIS_Tools_2025_Core.BIS_Log;
using static IC_Loader_Pro.Module1;
using IC_Rules_2025;
using IC_Loader_Pro.Services;
using IC_Loader_Pro.Models;
using BIS_Tools_DataModels_2025;

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

            try
            {
                // This work will be done on a background thread to keep the UI responsive.
                var summaryList = await QueuedTask.Run(static () =>
                {
                    var outlookService = new OutlookService();
                    var summaries = new List<ICQueueSummary>();

                    // Get the list of IC Types to process from your rules engine.
                    List<String> icTypes = new List<string> { "CEA", "DNA"};//IcRules.ReturnIcTypes();

                    foreach (string icType in icTypes)
                    {
                        try
                        {
                            // Get the specific settings for this queue, including the folder name.
                            IcGisTypeSetting icSetting = IcRules.ReturnIcGisTypeSettings(icType);
                            string outlookFolderName = icSetting.EmailFolderSet.InboxFolderName;

                            // Call our service to get the detailed list of emails for this folder.
                            List<EmailItem> emailsInQueue = outlookService.GetEmailsFromSubfolder(outlookFolderName);

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
