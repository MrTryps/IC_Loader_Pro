using ArcGIS.Desktop.Framework;
using ArcGIS.Desktop.Framework.Threading.Tasks;
using BIS_Tools_DataModels_2025;
using IC_Loader_Pro.Models;
using IC_Loader_Pro.Services;
using IC_Rules_2025;
using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using static BIS_Log;
using static IC_Loader_Pro.Module1;
using EmailType = IC_Loader_Pro.Models.EmailType;
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
                    Log.RecordMessage($"Verification: _ListOfIcEmailTypeSummaries now contains {_ListOfIcEmailTypeSummaries.Count} items.", BisLogMessageType.Note);
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

            EmailType? userSelectedType = null;

            var unprocessedEmails = new List<UnprocessedEmailInfo>();

            // This is now our main processing loop.
            // It continues as long as the selected queue still has emails.
            while (_emailQueues.TryGetValue(SelectedIcType.Name, out var emailsToProcess) && emailsToProcess.Any())
            {
                // Get the top email from the list for this iteration.
                var currentEmailSummary = emailsToProcess.First();
                EmailItem emailToProcess = null; // To hold the full email object for cleanup

                try
                {
                    // --- 1. Fetch and Classify ---
                    var icSetting = IcRules.ReturnIcGisTypeSettings(SelectedIcType.Name);
                    string fullOutlookPath = icSetting.OutlookInboxFolderPath;
                    var (storeName, folderPath) = OutlookService.ParseOutlookPath(fullOutlookPath);
                    var outlookService = new OutlookService();
                    emailToProcess = await QueuedTask.Run(() => outlookService.GetEmailById(folderPath, currentEmailSummary.Emailid, storeName));

                    if (emailToProcess == null)
                    {
                        StatusMessage = "Error: Could not retrieve email. Skipping to next.";
                        unprocessedEmails.Add(new UnprocessedEmailInfo { Subject = currentEmailSummary.Subject, Reason = "Failed to retrieve from Outlook." });
                        emailsToProcess.RemoveAt(0); // Remove the failed email and continue
                        continue;
                    }

                    var classifier = new EmailClassifierService(IcRules, Log);
                    var classification = classifier.ClassifyEmail(emailToProcess);

                    // --- 2. Handle Manual Classification Pop-up ---
                    if (classification.Type == EmailType.Unknown || classification.Type == EmailType.EmptySubjectline)
                    {
                        // This logic correctly shows the pop-up and handles cancellation
                        // If the user cancels, we return from the method entirely, stopping the loop.
                        if (await RequestManualEmailClassification(emailToProcess) is EmailType selectedType)
                        {
                            userSelectedType = selectedType;
                        }
                        else
                        {
                            StatusMessage = "Processing canceled by user.";
                            unprocessedEmails.Add(new UnprocessedEmailInfo { Subject = emailToProcess.Subject, Reason = "Canceled during manual classification." });
                            continue; // Skip to next email
                        }
                    }

                    // --- 3. Update UI and Process in Background ---
                    UpdateEmailInfo(emailToProcess, classification);

                    var namedTests = new IcNamedTests(Log, PostGreTool);
                    var processingService = new EmailProcessingService(IcRules, namedTests, Log);
                    IcTestResult finalResult = await processingService.ProcessEmailAsync(emailToProcess, classification, SelectedIcType.Name, folderPath, storeName, userSelectedType);
                    if (!finalResult.Passed)
                    {
                        unprocessedEmails.Add(new UnprocessedEmailInfo
                        {
                            Subject = emailToProcess.Subject,
                            // Use the last comment as the reason, which will be accurate.
                            Reason = finalResult.Comments.LastOrDefault() ?? "An unknown error occurred."
                        });
                    }
                    // --- 4. Update Final Status and Stats ---
                    UpdateQueueStats(finalResult);
                }
                catch (Exception ex)
                {
                    StatusMessage = "An error occurred during processing.";
                    Log.RecordError($"Error processing email ID {currentEmailSummary.Emailid}", ex, "ProcessSelectedQueueAsync");
                    unprocessedEmails.Add(new UnprocessedEmailInfo { Subject = currentEmailSummary.Subject, Reason = "An unexpected error occurred." });
                }
                finally
                {
                    // --- 5. Cleanup and Loop ---
                    // ALWAYS remove the processed email from the queue.
                    emailsToProcess.RemoveAt(0);

                    // Clean up the temporary attachment folder.
                    CleanupTempFolder(emailToProcess);

                    // Update the email count in the UI
                    SelectedIcType.EmailCount = emailsToProcess.Count;
                }
            }

            if (unprocessedEmails.Any())
            {
                var summaryMessage = new System.Text.StringBuilder();
                summaryMessage.AppendLine("The following emails were not fully processed and may require manual attention:");
                summaryMessage.AppendLine();

                foreach (var info in unprocessedEmails)
                {
                    summaryMessage.AppendLine($"• {info.Subject}: {info.Reason}");
                }

                ArcGIS.Desktop.Framework.Dialogs.MessageBox.Show(summaryMessage.ToString(), "Queue Processing Summary");
            }

            // This message shows when the loop is finished and the queue is empty.
            StatusMessage = $"Queue '{SelectedIcType.Name}' is empty.";
        }

        private async Task<EmailType?> RequestManualEmailClassification(EmailItem email)
        {
            var attachmentNames = email.Attachments.Select(a => a.FileName).ToList();
            var popupViewModel = new ViewModels.ManualEmailClassificationViewModel(email.SenderEmailAddress, email.Subject, attachmentNames);
            var popupWindow = new Views.ManualEmailClassificationWindow
            {
                DataContext = popupViewModel,
                Owner = FrameworkApplication.Current.MainWindow
            };

            if (popupWindow.ShowDialog() == true)
            {                
                return popupViewModel.SelectedEmailType; // User clicked OK
            }
            return null; // User canceled
        }

        private void UpdateEmailInfo(EmailItem email, EmailClassificationResult classification)
        {
            CurrentEmailSubject = email.Subject;
            CurrentPrefId = classification.PrefIds.FirstOrDefault() ?? "N/A";
            CurrentAltId = classification.AltIds.FirstOrDefault() ?? "N/A";
            CurrentActivityNum = classification.ActivityNums.FirstOrDefault() ?? "N/A";
            CurrentDelId = "Pending";
            StatusMessage = "Processing...";
        }

        private void UpdateQueueStats(IcTestResult finalResult)
        {
            if (finalResult.Passed)
            {
                SelectedIcType.PassedCount++;
                StatusMessage = "Email processed successfully. Loading next...";
            }
            else
            {
                SelectedIcType.FailedCount++;
                StatusMessage = $"Processing failed: {string.Join(" ", finalResult.Comments)}. Loading next...";
            }
        }

        private void CleanupTempFolder(EmailItem email)
        {
            if (email != null && !string.IsNullOrEmpty(email.TempFolderPath))
            {
                try
                {
                    if (Directory.Exists(email.TempFolderPath))
                    {
                        Directory.Delete(email.TempFolderPath, true);
                    }
                }
                catch (Exception ex)
                {
                    Log.RecordError($"Failed to delete temp folder: {email.TempFolderPath}", ex, "CleanupTempFolder");
                }
            }
        }

    }
}