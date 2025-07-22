using ArcGIS.Desktop.Framework;
using ArcGIS.Desktop.Framework.Contracts;
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
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using static BIS_Log;
using static IC_Loader_Pro.Module1;
using EmailType = BIS_Tools_DataModels_2025.EmailType;
using Exception = System.Exception;
using Outlook = Microsoft.Office.Interop.Outlook;

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
                StatusMessage = "Connecting to Outlook and loading queues...";
            });

            Outlook.Application outlookApp = null; // Our single, shared instance

            Log.RecordMessage("Refreshing IC Queue summaries from source...", BisLogMessageType.Note);

            try
            {
                // --- Step 2: Background work ---
                outlookApp = new Outlook.Application();
                var outlookService = new OutlookService(); // For the responsiveness check
                                                           // Check for responsiveness first
                if (!outlookService.IsOutlookResponsive(outlookApp))
                {
                    throw new Services.OutlookNotResponsiveException("Outlook is not running or is not responsive.");
                }

                var result = await GetEmailSummariesAsync(outlookApp);
                _emailQueues = result.Success;

                if (result.FailedQueues.Any())
                {
                    string failedQueueNames = string.Join(", ", result.FailedQueues);
                    string errorMessage = $"The following email queues could not be loaded because their Outlook folders were not found:\n\n- {failedQueueNames}";
                    await RunOnUIThread(() =>
                    {
                        ArcGIS.Desktop.Framework.Dialogs.MessageBox.Show(errorMessage, "Missing Outlook Folders");
                    });
                }


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
                Log.RecordError("Could not connect to Outlook.", ex, nameof(RefreshICQueuesAsync));
                await RunOnUIThread(async () =>
                {
                    await OutlookService.TryRestartOutlook();
                    StatusMessage = "Outlook restart attempted. Please try refreshing again.";
                });

                // After the restart attempt, you might want to try the refresh again or just inform the user.
                StatusMessage = "Outlook restart attempted. Please try refreshing the queues again.";

                //await RunOnUIThread(() =>
                //{
                //    ArcGIS.Desktop.Framework.Dialogs.MessageBox.Show(
                //        $"{ex.Message}\nPlease ensure Outlook is open and running correctly before refreshing.",
                //        "Outlook Connection Error",
                //        System.Windows.MessageBoxButton.OK,
                //        System.Windows.MessageBoxImage.Warning);
                //    StatusMessage = "Could not connect to Outlook.";
                //});
            }           
            catch (Exception ex)
            {
                Log.RecordError("A fatal error occurred while refreshing the IC Queues.", ex, nameof(RefreshICQueuesAsync));
                await RunOnUIThread(() => { StatusMessage = "Error loading email queues."; });
            }
            finally
            {
                // --- Step 4: Re-enable UI ---
                if (outlookApp != null)
                {
                    Marshal.ReleaseComObject(outlookApp);
                }
                await RunOnUIThread(() => { IsUIEnabled = true; });
            }
        }

        /// <summary>
        /// This helper method is now fully async from top to bottom.
        /// </summary>
        private async Task<(Dictionary<string, List<EmailItem>> Success, List<string> FailedQueues)> GetEmailSummariesAsync(Outlook.Application outlookApp)
        {
            var rulesEngine = Module1.IcRules;
            var queues = new Dictionary<string, List<EmailItem>>(StringComparer.OrdinalIgnoreCase);
            var failedQueues = new List<string>(); // List to track failures

            // Determine which service to use
            GraphApiService graphService = null;
            var outlookService = new OutlookService();

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
                            outlookService.GetEmailsFromFolderPath(outlookApp,outlookFolderPath, testSender, testModeFlag));
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
                    Log.RecordError($"An error occurred while processing queue '{icType}'. It will be skipped.", ex, nameof(GetEmailSummariesAsync));
                    failedQueues.Add(icType);
                }
            }
            return (queues, failedQueues);
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

            // NEW CHANGE: Initialize the named tests service once at the start.
            // This is safer and prevents re-creating it inside a try-catch block.
            var namedTests = new IcNamedTests(Log, PostGreTool);
            var currentEmailSummary = emailsToProcess.First();
            EmailItem emailToProcess = null;
            bool wasAutoAdvanced = false; // Flag to track if we should auto-advance
            Outlook.Application outlookApp = null;

            try
            {
                outlookApp = new Outlook.Application();
                // --- 2. Fetch and Classify the Email ---
                var icSetting = IcRules.ReturnIcGisTypeSettings(SelectedIcType.Name);
                var (storeName, folderPath) = OutlookService.ParseOutlookPath(icSetting.OutlookInboxFolderPath);
                var outlookService = new OutlookService();
                emailToProcess = await QueuedTask.Run(() => outlookService.GetEmailById(outlookApp,folderPath, currentEmailSummary.Emailid, storeName));

                if (emailToProcess == null)
                {
                    ArcGIS.Desktop.Framework.Dialogs.MessageBox.Show($"Could not retrieve the email: '{currentEmailSummary.Subject}'. It will be skipped.", "Email Retrieval Error");
                    await ProcessNextEmail();
                    return;
                }

                var classifier = new EmailClassifierService(IcRules, Log);
                var classification = classifier.ClassifyEmail(emailToProcess);

                bool userDidSelect = false;
                EmailType finalEmailType = classification.Type;
                //EmailType? userSelectedType = null;
                if (classification.Type == EmailType.Unknown || classification.Type == EmailType.EmptySubjectline)
                {
                    var (wasSelected, selectedType) = await RequestManualEmailClassification(emailToProcess);
                    if (wasSelected)
                    {
                        userDidSelect = true;
                        finalEmailType = selectedType;                        
                    }
                    else
                    {
                        await ProcessNextEmail(); // User canceled, skip to next.
                        return;
                    }
                }

                UpdateEmailInfo(emailToProcess, classification, userDidSelect, finalEmailType);

                // --- 3. Process the Email and Handle the Result ---
                var processingService = new EmailProcessingService(IcRules, namedTests, Log);
                EmailProcessingResult processingResult = await processingService.ProcessEmailAsync(outlookApp,emailToProcess, classification, SelectedIcType.Name, folderPath, storeName, userDidSelect, finalEmailType);

                _currentEmailTestResult = processingResult.TestResult;
                UpdateQueueStats(_currentEmailTestResult);

                if (!_currentEmailTestResult.Passed)
                {
                    wasAutoAdvanced = true; // Mark for auto-advancement
                    ShowTestResultWindow(_currentEmailTestResult);
                    await ProcessNextEmail(); // Auto-fail: show results and advance
                    return;
                }

                // --- 4. On Success, Populate UI and Wait for User Input ---
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
                IsEmailActionEnabled = true; // Enable Save/Skip/Reject buttons
            }
            catch (Exception ex)
            {
                Log.RecordError($"An unexpected error occurred while processing email ID {currentEmailSummary.Emailid}", ex, "ProcessSelectedQueueAsync");
                ArcGIS.Desktop.Framework.Dialogs.MessageBox.Show($"An unexpected error occurred. The application will advance to the next email.", "Processing Error");
                await ProcessNextEmail(); // On unexpected error, advance to the next email.
            }
            finally
            {
                // NEW CHANGE: The 'finally' block is now the single, guaranteed place
                // where the processed email is removed from the queue and cleaned up.
                // This prevents all the bugs related to double-removal or getting stuck.
                if (emailsToProcess.Any() && emailsToProcess.First() == currentEmailSummary)
                {
                    emailsToProcess.RemoveAt(0);
                }
                CleanupTempFolder(emailToProcess);
                if (SelectedIcType != null)
                {
                    SelectedIcType.EmailCount = emailsToProcess.Count;
                }
                // If the email was auto-failed or had an error, immediately process the next one.
                if (wasAutoAdvanced)
                {
                    await ProcessNextEmail();
                }
                //if (outlookApp != null)
                //{
                //    Marshal.ReleaseComObject(outlookApp);
                //}
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

        private async Task<(bool wasSelected, EmailType selectedType)> RequestManualEmailClassification(EmailItem email)
        {
            var attachmentNames = email.Attachments.Select(a => a.FileName).ToList();
            // Ensure you have a ViewModel for your popup, e.g., ManualEmailClassificationViewModel
            var popupViewModel = new ViewModels.ManualEmailClassificationViewModel(email.SenderEmailAddress, email.Subject, attachmentNames);
            var popupWindow = new Views.ManualEmailClassificationWindow
            {
                DataContext = popupViewModel,
                Owner = FrameworkApplication.Current.MainWindow
            };

            if (popupWindow.ShowDialog() == true)
            {
                // Return true and the user's selection
                return (true, popupViewModel.SelectedEmailType);
            }
            // Return false and a default value
            return (false, EmailType.Unknown);
        }

        private void UpdateEmailInfo(EmailItem email, EmailClassificationResult classification, bool wasManuallySelected, EmailType finalType)
        {
            CurrentEmailId = email.Emailid;
            CurrentEmailSubject = email.Subject;
            CurrentPrefId = classification.PrefIds.FirstOrDefault() ?? "N/A";
            CurrentAltId = classification.AltIds.FirstOrDefault() ?? "N/A";
            CurrentActivityNum = classification.ActivityNums.FirstOrDefault() ?? "N/A";
            CurrentDelId = "Pending";

            // If a type was manually selected, use that for the status message.
            if (wasManuallySelected)
            {
                StatusMessage = $"Processing email as type '{finalType.Value}'...";
            }
            else
            {
                StatusMessage = "Processing...";
            }
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

        /// <summary>
        /// A temporary diagnostic method to debug Outlook folder access issues.
        /// </summary>
        /// <summary>
        /// A temporary diagnostic method that calls the folder consistency test in the service layer.
        /// </summary>
        private async Task TestFolderAccessAsync2()
        {
            StatusMessage = "Running folder consistency test...";


            // We need an instance of the service to call the method.
            var namedTests = new IcNamedTests(Log, PostGreTool);
            var processingService = new EmailProcessingService(IcRules, namedTests, Log);

            // Call the test method.
            await processingService.RunFolderConsistencyTestAsync();

            StatusMessage = "Consistency test complete. Check the log file.";
        }


        /// <summary>
        /// A temporary diagnostic method to debug Outlook folder access issues.
        /// </summary>
        private async Task TestFolderAccessAsyncB()
        {
            Log.AddBlankLine();
            Log.RecordMessage("--- Starting Outlook Folder Diagnostic Test ---", BisLogMessageType.Note);
            StatusMessage = "Running folder diagnostic...";

            // --- CONFIGURE YOUR TEST HERE ---
            // Change these values to test different paths
            string targetStoreName = "DEP srpgis_cea [DEP]";
            string targetFolderPath = "Inbox\\CEA_Processed";
            // --------------------------------

            await QueuedTask.Run(() =>
            {
                Outlook.Application outlookApp = null;
                Outlook.NameSpace mapiNamespace = null;
                Outlook.Store targetStore = null;
                Outlook.MAPIFolder currentFolder = null;

                try
                {
                    outlookApp = new Outlook.Application();
                    mapiNamespace = outlookApp.GetNamespace("MAPI");

                    // 1. List all available stores
                    Log.RecordMessage("--- Available Mailbox Stores ---", BisLogMessageType.Note);
                    foreach (Outlook.Store store in mapiNamespace.Stores)
                    {
                        Log.RecordMessage($"Found Store: '{store.DisplayName}'", BisLogMessageType.Note);
                        Marshal.ReleaseComObject(store);
                    }
                    Log.RecordMessage("---------------------------------", BisLogMessageType.Note);

                    // 2. Try to find the target store
                    targetStore = mapiNamespace.Stores
                        .Cast<Outlook.Store>()
                        .FirstOrDefault(s => s.DisplayName.Equals(targetStoreName, StringComparison.OrdinalIgnoreCase));

                    if (targetStore == null)
                    {
                        Log.RecordError($"TEST FAILED: Could not find the store named '{targetStoreName}'. Please check for typos or if the mailbox is added to Outlook.", null, "TestFolderAccess");
                        return;
                    }
                    Log.RecordMessage($"Successfully found store: '{targetStore.DisplayName}'", BisLogMessageType.Note);

                    // 3. Traverse the folder path step-by-step
                    currentFolder = targetStore.GetRootFolder();
                    Log.RecordMessage($"Starting search from root folder: '{currentFolder.Name}'", BisLogMessageType.Note);

                    var folderNames = targetFolderPath.Split('\\');
                    foreach (var name in folderNames)
                    {
                        Outlook.MAPIFolder nextFolder = null;
                        try
                        {
                            nextFolder = currentFolder.Folders[name];
                            Log.RecordMessage($"  -> Successfully entered subfolder: '{name}'", BisLogMessageType.Note);

                            // Release the previous folder and move to the next one
                            if (currentFolder != targetStore.GetRootFolder()) Marshal.ReleaseComObject(currentFolder);
                            currentFolder = nextFolder;
                        }
                        catch
                        {
                            Log.RecordError($"TEST FAILED: Could not find the subfolder named '{name}' inside of '{currentFolder.Name}'.", null, "TestFolderAccess");
                            if (nextFolder != null) Marshal.ReleaseComObject(nextFolder);
                            return; // Stop the test
                        }
                    }

                    Log.RecordMessage($"--- DIAGNOSTIC SUCCEEDED: Successfully navigated to the final folder '{currentFolder.FolderPath}' ---", BisLogMessageType.Note);
                }
                catch (Exception ex)
                {
                    Log.RecordError("An unexpected exception occurred during the diagnostic test.", ex, "TestFolderAccess");
                }
                finally
                {
                    // Clean up all COM objects
                    if (currentFolder != null) Marshal.ReleaseComObject(currentFolder);
                    if (targetStore != null) Marshal.ReleaseComObject(targetStore);
                    if (mapiNamespace != null) Marshal.ReleaseComObject(mapiNamespace);
                    if (outlookApp != null) Marshal.ReleaseComObject(outlookApp);
                }
            });

            StatusMessage = "Diagnostic test complete. Check the log file.";
        }     

    }
}