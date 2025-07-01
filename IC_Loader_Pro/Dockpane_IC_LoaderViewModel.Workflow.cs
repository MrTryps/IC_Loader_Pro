using ArcGIS.Desktop.Framework;
using ArcGIS.Desktop.Framework.Threading.Tasks;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using static IC_Loader_Pro.Module1;
using IC_Loader_Pro.Services;
using IC_Loader_Pro.Models;
using BIS_Tools_DataModels_2025;
using static BIS_Log;

namespace IC_Loader_Pro
{
    internal partial class Dockpane_IC_LoaderViewModel
    {
        // --- MASTER SWITCH ---
        private const bool useGraphApi = false;

        private async Task RefreshICQueuesAsync()
        {
            // --- Step 1: Initial UI update ---
            Log.RecordMessage("Step 1: Calling RunOnUIThread to disable UI.", BisLogMessageType.Note);
            await RunOnUIThread(() =>
            {
                IsUIEnabled = false;
                StatusMessage = "Loading email queues...";
            });
            Log.RecordMessage("Step 1: Completed.", BisLogMessageType.Note);

            Log.RecordMessage("Refreshing IC Queue summaries from source...", BisLogMessageType.Note);

            try
            {
                // --- Step 2: Background work ---
                List<ICQueueSummary> summaryList = await GetEmailSummariesAsync();
                Log.RecordMessage($"Step 2: Background work complete. Found {summaryList.Count} summaries.", BisLogMessageType.Note);

                // --- Step 3: Final UI update ---
                Log.RecordMessage("Step 3: Calling RunOnUIThread to update UI with results.", BisLogMessageType.Note);
                await RunOnUIThread(() =>
                {
                    lock (_lockQueueCollection)
                    {
                        _ListOfIcEmailTypeSummaries.Clear();
                        foreach (var summary in summaryList)
                        {
                            _ListOfIcEmailTypeSummaries.Add(summary);
                        }
                    }

                    SelectedIcType = PublicListOfIcEmailTypeSummaries.FirstOrDefault();
                    Log.RecordMessage($"Successfully loadedxxx {PublicListOfIcEmailTypeSummaries.Count} queues.", BisLogMessageType.Note);

                    if (SelectedIcType != null)
                    {
                        StatusMessage = $"Ready. Default queue '{SelectedIcType.Name}' selected.";
                    }
                    else
                    {
                        StatusMessage = "No emails found in the specified queues.";
                    }
                });
                Log.RecordMessage("Step 3: Completed.", BisLogMessageType.Note);
            }
            catch (Exception ex)
            {
                Log.RecordError("A fatal error occurred while refreshing the IC Queues.", ex, nameof(RefreshICQueuesAsync));
                await RunOnUIThread(() => { StatusMessage = "Error loading email queues."; });
            }
            finally
            {
                // --- Step 4: Re-enable UI ---
                Log.RecordMessage("Step 4: Calling RunOnUIThread to re-enable UI.", BisLogMessageType.Note);
                await RunOnUIThread(() => { IsUIEnabled = true; });
                Log.RecordMessage("Step 4: Completed.", BisLogMessageType.Note);
            }
        }

        /// <summary>
        /// This helper method is now fully async from top to bottom.
        /// </summary>
        private async Task<List<ICQueueSummary>> GetEmailSummariesAsync()
        {
            var rulesEngine = Module1.IcRules;
            var summaries = new List<ICQueueSummary>();

            // Determine which service to use
            GraphApiService graphService = null;
            OutlookService outlookService = null;

            if (useGraphApi)
            {
                Log.RecordMessage("Using Microsoft Graph API Service.", BisLogMessageType.Note);
                graphService = await GraphApiService.CreateAsync();
            }
            else
            {
                Log.RecordMessage("Using Outlook Interop Service.", BisLogMessageType.Note);
                outlookService = new OutlookService();
            }

            foreach (string icType in rulesEngine.ReturnIcTypes())
            {
                try
                {
                    IcGisTypeSetting icSetting = rulesEngine.ReturnIcGisTypeSettings(icType);
                    string outlookFolderPath = icSetting.OutlookInboxFolderPath;
                    string testSender = icSetting.TestSenderEmail;
                    // --- LOCAL TEST FLAG ---
                    // Set the test mode directly in the code.
                    // true  = Filter FOR emails from the test sender only.
                    // false = Filter OUT emails from the test sender.
                    // null  = Disable test filtering.
                    bool? testModeFlag = true;

                    if (string.IsNullOrEmpty(outlookFolderPath))
                    {
                        Log.RecordMessage($"Skipping queue '{icType}' because OutlookFolderPath is not configured.", BisLogMessageType.Warning);
                        continue;
                    }

                    List<EmailItem> emailsInQueue;
                    if (useGraphApi)
                    {
                        // Await the async Graph call directly. No .Result.
                        emailsInQueue = await graphService.GetEmailsFromFolderPathAsync(outlookFolderPath, testSender, testModeFlag);
                    }
                    else
                    {
                        // Use QueuedTask.Run to move the synchronous Outlook Interop call off the UI thread.
                        emailsInQueue = await QueuedTask.Run(() =>
                            outlookService.GetEmailsFromFolderPath(outlookFolderPath, testSender, testModeFlag));
                    }

                    summaries.Add(new ICQueueSummary
                    {
                        Name = icType,
                        EmailCount = emailsInQueue.Count,
                        PassedCount = 0,
                        SkippedCount = 0,
                        FailedCount = 0
                    });
                }
                catch (Exception ex)
                {
                    Log.RecordError($"An error occurred while processing queue '{icType}'.", ex, nameof(GetEmailSummariesAsync));
                }
            }
            return summaries;
        }
    }
}