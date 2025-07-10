using ArcGIS.Desktop.Framework;
using ArcGIS.Desktop.Framework.Threading.Tasks;
using BIS_Tools_DataModels_2025;
using IC_Loader_Pro.Models;
using IC_Loader_Pro.Services;
using IC_Rules_2025;
using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using static BIS_Log;
using static IC_Loader_Pro.Module1;
using Exception = System.Exception;

namespace IC_Loader_Pro
{
    internal partial class Dockpane_IC_LoaderViewModel
    {
        // --- MASTER SWITCH ---
        private const bool useGraphApi = false;
        private Dictionary<string, List<EmailItem>> _emailQueues;

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
                _emailQueues = await GetEmailSummariesAsync();
                // Populate the UI summary list from the full data.
                var summaryList = _emailQueues.Select(kvp => new ICQueueSummary
                {
                    Name = kvp.Key,
                    EmailCount = kvp.Value.Count
                }).ToList();
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
        private async Task<Dictionary<string, List<EmailItem>>> GetEmailSummariesAsync()
        {
            var rulesEngine = Module1.IcRules;
            var queues = new Dictionary<string, List<EmailItem>>(StringComparer.OrdinalIgnoreCase);

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

                    queues[icType] = emailsInQueue;
                }
                catch (Exception ex)
                {
                    Log.RecordError($"An error occurred while processing queue '{icType}'.", ex, nameof(GetEmailSummariesAsync));
                }
            }
            return queues;
        }

        /// <summary>
        /// Kicks off the processing for the currently selected IC queue.
        /// </summary>
        private async Task ProcessSelectedQueueAsync()
        {                      
            if (SelectedIcType == null)
            {
                StatusMessage = "No queue selected.";
                return;
            }

            // Find the list of emails for the selected queue
            if (!_emailQueues.TryGetValue(SelectedIcType.Name, out var emailsToProcess) || !emailsToProcess.Any())
            {
                StatusMessage = $"Queue '{SelectedIcType.Name}' is empty.";
                return;
            }

            // Get the first email from the list
            var firstEmail = emailsToProcess.First();
            var icSetting = IcRules.ReturnIcGisTypeSettings(SelectedIcType.Name);
            StatusMessage = $"Loading email: {firstEmail.Subject}...";
            string fullOutlookPath = icSetting.OutlookInboxFolderPath;
            var (storeName, folderPath) = OutlookService.ParseOutlookPath(fullOutlookPath);

            try
            {
                // Instantiate the services needed for processing
                var namedTests = new IcNamedTests(Log, PostGreTool);
                var classifier = new EmailClassifierService(IcRules, Log);
                var attachmentService = new AttachmentService(IcRules, namedTests,FileTool, Log);

                // This is the main orchestrator for processing a single email
                var processingService = new EmailProcessingService(IcRules, namedTests,Log);

                // Process the single email
                IcTestResult finalResult = await processingService.ProcessEmailAsync(firstEmail.Emailid, folderPath, storeName);

                // Update the UI with the results of the processing
                if (finalResult.Passed)
                {
                    StatusMessage = "Email processed successfully.";
                }
                else
                {
                    // Join all comments from the test result hierarchy for a detailed status
                    var allComments = finalResult.Comments.Concat(finalResult.SubTestResults.SelectMany(sr => sr.Comments));
                    StatusMessage = $"Processing failed: {string.Join(" ", allComments)}";
                }
            }
            catch (Exception ex)
            {
                StatusMessage = "An error occurred while processing the email.";
                Log.RecordError($"Error during ProcessSelectedQueueAsync for email ID {firstEmail.Emailid}", ex, "ProcessSelectedQueueAsync");
            }
        }



    }
}