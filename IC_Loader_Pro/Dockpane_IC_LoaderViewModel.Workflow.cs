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
                int totalEmailCount = _emailQueues.Values.Sum(emailList => emailList.Count);
                var summaryList = _emailQueues.Select(kvp => new ICQueueSummary
                {
                    Name = kvp.Key,
                    EmailCount = kvp.Value.Count
                }).ToList();
                Log.RecordMessage($"Step 2: Background work complete. Found {totalEmailCount} summaries.", BisLogMessageType.Note);

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
            catch (OutlookNotResponsiveException ex)
            {
                await RunOnUIThread(() =>
                {
                    ArcGIS.Desktop.Framework.Dialogs.MessageBox.Show(
                        $"{ex.Message}\nPlease ensure Outlook is open and running correctly before refreshing.",
                        "Outlook Connection Error",
                        System.Windows.MessageBoxButton.OK,
                        System.Windows.MessageBoxImage.Warning);
                    StatusMessage = "Could not connect to Outlook.";
                });
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
                catch (OutlookNotResponsiveException ex)
                {
                    Log.RecordError("Could not connect to Outlook.", ex, nameof(GetEmailSummariesAsync));
                    // We only need to show the message once, so we'll re-throw to stop the loop.
                    throw;
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
            // --- 1. Initial UI and Configuration Setup ---
            IsEmailActionEnabled = false;
            _foundFileSets.Clear();

            if (SelectedIcType == null || !_emailQueues.TryGetValue(SelectedIcType.Name, out var emailsToProcess) || !emailsToProcess.Any())
            {
                CurrentEmailSubject = "Queue is empty.";
                StatusMessage = $"Queue '{SelectedIcType?.Name}' is empty.";
                return;
            }

            // Check for IC Rules settings first. A failure here is a fatal configuration problem.
            var icSetting = IcRules.ReturnIcGisTypeSettings(SelectedIcType.Name);
            if (icSetting == null)
            {
                var errorMsg = $"Configuration settings for '{SelectedIcType.Name}' not found. Cannot proceed.";
                Log.RecordError(errorMsg, null, nameof(ProcessSelectedQueueAsync));
                ArcGIS.Desktop.Framework.Dialogs.MessageBox.Show(errorMsg, "Configuration Error");
                return;
            }

            // This try-catch is ONLY for initialization. A failure here is also a fatal configuration error.
            IcNamedTests namedTests;
            try
            {
                namedTests = new IcNamedTests(Log, PostGreTool);
            }
            catch (Exception ex)
            {
                Log.RecordError("Fatal error: Could not initialize the Named Tests service. A required test rule is likely missing from the database.", ex, nameof(ProcessSelectedQueueAsync));
                ArcGIS.Desktop.Framework.Dialogs.MessageBox.Show("Could not start processing. A required test rule is missing. Please check the logs.", "Configuration Error");
                return;
            }


            // --- 2. Main Email Processing Block ---
            var currentEmailSummary = emailsToProcess.First();
            EmailItem emailToProcess = null;

            // This try-catch block is now only responsible for errors related to this specific email.
            try
            {
                var (storeName, folderPath) = OutlookService.ParseOutlookPath(icSetting.OutlookInboxFolderPath);
                var outlookService = new OutlookService();
                emailToProcess = await QueuedTask.Run(() => outlookService.GetEmailById(folderPath, currentEmailSummary.Emailid, storeName));

                if (emailToProcess == null)
                {
                    ArcGIS.Desktop.Framework.Dialogs.MessageBox.Show($"Could not retrieve the email with subject: '{currentEmailSummary.Subject}'. It will be skipped.", "Email Retrieval Error");
                    await ProcessNextEmail(); // The finally block will handle cleanup
                    return;
                }

                var classifier = new EmailClassifierService(IcRules, Log);
                var classification = classifier.ClassifyEmail(emailToProcess);

                EmailType? userSelectedType = null;
                if (classification.Type == EmailType.Unknown || classification.Type == EmailType.EmptySubjectline)
                {
                    if (await RequestManualEmailClassification(emailToProcess) is EmailType selectedType)
                    {
                        userSelectedType = selectedType;
                    }
                    else
                    {
                        // User canceled the pop-up
                        await ProcessNextEmail();
                        return;
                    }
                }

                UpdateEmailInfo(emailToProcess, classification);

                var processingService = new EmailProcessingService(IcRules, namedTests, Log);
                EmailProcessingResult processingResult = await processingService.ProcessEmailAsync(emailToProcess, classification, SelectedIcType.Name, folderPath, storeName, userSelectedType);

                _currentEmailTestResult = processingResult.TestResult;
                _currentAttachmentAnalysis = processingResult.AttachmentAnalysis;
                UpdateQueueStats(_currentEmailTestResult);

                if (!_currentEmailTestResult.Passed)
                {
                    SelectedIcType.FailedCount++;
                    StatusMessage = $"Auto-fail: {_currentEmailTestResult.Comments.LastOrDefault()}";
                    ShowTestResultWindow(_currentEmailTestResult);
                    // 1. Remove the failed email from the queue immediately.
                    emailsToProcess.RemoveAt(0);
                    await ProcessNextEmail();
                    return;
                }

                if (processingResult.AttachmentAnalysis?.IdentifiedFileSets?.Any() == true)
                {
                    await RunOnUIThread(() =>
                    {
                        _foundFileSets.Clear();
                        foreach (var fs in processingResult.AttachmentAnalysis.IdentifiedFileSets)
                        {
                            _foundFileSets.Add(new ViewModels.FileSetViewModel(fs));
                        }
                    });
                }

                StatusMessage = "Ready for review.";
                IsEmailActionEnabled = true;
            }
            catch (Exception ex)
            {
                Log.RecordError($"An unexpected error occurred while processing email ID {currentEmailSummary.Emailid}", ex, "ProcessSelectedQueueAsync");
                ArcGIS.Desktop.Framework.Dialogs.MessageBox.Show($"An unexpected error occurred while processing '{currentEmailSummary.Subject}'. The application will advance to the next email.", "Processing Error");
                emailsToProcess.RemoveAt(0);
                await ProcessNextEmail();
            }
            finally
            {
                //// This cleanup block is now simpler and always runs for the processed email.
                //if (emailsToProcess.Any() && emailsToProcess.First() == currentEmailSummary)
                //{
                //    emailsToProcess.RemoveAt(0);
                //}
                CleanupTempFolder(emailToProcess);
                if (SelectedIcType != null)
                {
                    SelectedIcType.EmailCount = emailsToProcess.Count;
                }
            }
        }

        /// <summary>
        /// A helper method that simply calls the main processing logic.
        /// This will be triggered by the user action buttons.
        /// </summary>
        private async Task ProcessNextEmail()
        {
            await ProcessSelectedQueueAsync();
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
            CurrentEmailId = email.Emailid;
            CurrentEmailSubject = email.Subject;
            CurrentPrefId = classification.PrefIds.FirstOrDefault() ?? "N/A";
            CurrentAltId = classification.AltIds.FirstOrDefault() ?? "N/A";
            CurrentActivityNum = classification.ActivityNums.FirstOrDefault() ?? "N/A";
            CurrentDelId = "Pending";
            StatusMessage = "Processing...";
        }

        private void UpdateQueueStats(IcTestResult finalResult)
        {
            // The IcTestResult class aggregates the most severe action from all sub-tests.
            // We can check this final, cumulative action.
            switch (finalResult.CumulativeAction.ResultAction)
            {
                //case TestActionResponse.Pass:
                //    SelectedIcType.PassedCount++;
                //    StatusMessage = "Email processed successfully. Ready for review.";
                //    break;

                case TestActionResponse.Note:
                    // This is our new "Skip" condition, based on the test rule's action.
                    SelectedIcType.SkippedCount++;
                    StatusMessage = "Email skipped. Loading next...";
                    break;

                case TestActionResponse.Manual:
                case TestActionResponse.Fail:
                    // All other non-passing actions are considered failures.
                    SelectedIcType.FailedCount++;
                    StatusMessage = $"Processing failed: {string.Join(" ", finalResult.Comments)}. Please review.";
                    break;
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