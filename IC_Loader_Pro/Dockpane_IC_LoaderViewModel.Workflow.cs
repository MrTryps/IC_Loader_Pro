using ArcGIS.Core.CIM;
using ArcGIS.Core.Geometry;
using ArcGIS.Core.Internal.Geometry;
using ArcGIS.Desktop.Core;
using ArcGIS.Desktop.Framework;
using ArcGIS.Desktop.Framework.Threading.Tasks;
using ArcGIS.Desktop.Layouts;
using ArcGIS.Desktop.Mapping;
using BIS_Tools_DataModels_2025;
using IC_Loader_Pro.Helpers;
using IC_Loader_Pro.Models;
using IC_Loader_Pro.Services;
using IC_Loader_Pro.ViewModels;
using IC_Loader_Pro.Views;
using IC_Rules_2025;
using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Windows.Input;
using static BIS_Log;
using static IC_Loader_Pro.Module1;
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
            //Log.RecordMessage("Step 1: Calling RunOnUIThread to disable UI.", BisLogMessageType.Note);
            await RunOnUIThread(() =>
            {
                IsUIEnabled = false;
                StatusMessage = "Loading email queues...";
            });

            Outlook.Application outlookApp = null;

            try
            {
                outlookApp = new Outlook.Application();
                _emailQueues = await GetEmailSummariesAsync(outlookApp);
                // Populate the UI summary list from the full data.
                int totalEmailCount = _emailQueues.Values.Sum(emailList => emailList.Count);
                var summaryList = _emailQueues.Select(kvp => new ICQueueSummary
                {
                    Name = kvp.Key,
                    EmailCount = kvp.Value.Count
                }).ToList();
                Log.RecordMessage($"Found {totalEmailCount} IC Emails.", BisLogMessageType.Note);

                // --- Step 3: Final UI update ---
               // Log.RecordMessage("Step 3: Calling RunOnUIThread to update UI with results.", BisLogMessageType.Note);
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
                    //Log.RecordMessage($"Verification: _ListOfIcEmailTypeSummaries now contains {_ListOfIcEmailTypeSummaries.Count} items.", BisLogMessageType.Note);
                    SelectedIcType = PublicListOfIcEmailTypeSummaries.FirstOrDefault();
                    //Log.RecordMessage($"Successfully loaded {PublicListOfIcEmailTypeSummaries.Count} queues.", BisLogMessageType.Note);

                    if (SelectedIcType != null)
                    {
                        StatusMessage = $"Ready. Default queue '{SelectedIcType.Name}' selected.";
                    }
                    else
                    {
                        StatusMessage = "No emails found in the specified queues.";
                    }
                });
                //Log.RecordMessage("Step 3: Completed.", BisLogMessageType.Note);
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
                if (outlookApp != null) Marshal.ReleaseComObject(outlookApp);
               // Log.RecordMessage("Step 4: Calling RunOnUIThread to re-enable UI.", BisLogMessageType.Note);
                await RunOnUIThread(() => { IsUIEnabled = true; });
               // Log.RecordMessage("Step 4: Completed.", BisLogMessageType.Note);
            }
        }

        /// <summary>
        /// This helper method is now fully async from top to bottom.
        /// </summary>
        private async Task<Dictionary<string, List<EmailItem>>> GetEmailSummariesAsync(Outlook.Application outlookApp)
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
                    
                    if (string.IsNullOrEmpty(outlookFolderPath))
                    {
                        Log.RecordMessage($"Skipping queue '{icType}' because OutlookFolderPath is not configured.", BisLogMessageType.Warning);
                        continue;
                    }

                    List<EmailItem> emailsInQueue;
                    if (useGraphApi)
                    {
                        // Await the async Graph call directly. No .Result.
                        emailsInQueue = await graphService.GetEmailsFromFolderPathAsync(outlookApp,outlookFolderPath, testSender, Module1.IsInTestMode);
                    }
                    else
                    {
                        // Use QueuedTask.Run to move the synchronous Outlook Interop call off the UI thread.
                        emailsInQueue = await QueuedTask.Run(() =>
                            outlookService.GetEmailsFromFolderPath(outlookApp,outlookFolderPath, testSender, Module1.IsInTestMode));
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
        /// 
        private async Task ProcessSelectedQueueAsync()
        {
            // --- 1. Initial UI and Configuration Setup ---
            await PerformCleanupAsync();
            await ClearManuallyLoadedLayersAsync();
            IsEmailActionEnabled = false;
            _foundFileSets.Clear();
            _allProcessedShapes.Clear();

            if (SelectedIcType == null || !_emailQueues.TryGetValue(SelectedIcType.Name, out var emailsToProcess) || !emailsToProcess.Any())
            {
                CurrentEmailSubject = "Queue is empty.";
                StatusMessage = $"Queue '{SelectedIcType?.Name}' is empty.";
                return;
            }

            Outlook.Application outlookApp = null;
            var namedTests = new IcNamedTests(Log, PostGreTool);
            var currentEmailSummary = emailsToProcess.First();
            EmailItem emailToProcess = null;

            bool shouldAutoAdvance = false;

            try
            {
                outlookApp = new Outlook.Application();
                var (storeName, folderPath) = OutlookService.ParseOutlookPath(_currentIcSetting.OutlookInboxFolderPath);

                emailToProcess = await QueuedTask.Run(() => new OutlookService().GetEmailById(outlookApp, folderPath, currentEmailSummary.Emailid, storeName));
                _currentEmail = emailToProcess;

                if (emailToProcess == null)
                {
                    shouldAutoAdvance = true;
                    ArcGIS.Desktop.Framework.Dialogs.MessageBox.Show($"Could not retrieve the email: '{currentEmailSummary.Subject}'. It will be skipped.", "Email Retrieval Error");
                    return;
                }

                var classification = new EmailClassifierService(IcRules, Log).ClassifyEmail(emailToProcess);
                _currentClassification = classification;

                EmailType finalEmailType = classification.Type;
                if (classification.Type == EmailType.Unknown || classification.Type == EmailType.EmptySubjectline)
                {
                    var (wasSelected, selectedType) = await RequestManualEmailClassification(emailToProcess);
                    if (wasSelected)
                    {
                        finalEmailType = selectedType;
                        classification.WasManuallyClassified = true;
                    }
                    else
                    {
                        shouldAutoAdvance = true;
                        return;
                    }
                }

                UpdateEmailInfo(emailToProcess, classification, classification.WasManuallyClassified, finalEmailType);

                _currentSiteLocation = await GetSiteCoordinatesFromNjemsAsync(CurrentPrefId);

                var processingService = new EmailProcessingService(IcRules, namedTests, Log);
                EmailProcessingResult processingResult = await processingService.ProcessEmailAsync(outlookApp, emailToProcess, classification, SelectedIcType.Name, folderPath, storeName, classification.WasManuallyClassified, finalEmailType, GetSiteCoordinatesFromPostgreAsync);

                _currentEmailTestResult = processingResult.TestResult;
                _currentAttachmentAnalysis = processingResult.AttachmentAnalysis;

                if (processingResult.TestResult == null)
                {
                    shouldAutoAdvance = true;
                    return;
                }

                // Check for the signal from the processing service.
                if (processingResult.RequiresNoGisFilesDecision)
                {
                    // Create and show our new custom dialog window.
                    var dialogViewModel = new NoGisFilesViewModel(emailToProcess, _currentIcSetting.OutlookInboxFolderPath, outlookApp);
                    var dialog = new NoGisFilesWindow
                    {
                        DataContext = dialogViewModel,
                        Owner = FrameworkApplication.Current.MainWindow
                    };

                    dialog.ShowDialog(); // The workflow pauses here until the user makes a choice.

                    // Handle the user's choice from the dialog.
                    switch (dialog.Result)
                    {
                        case NoGisFilesWindow.UserChoice.Correspondence:
                            StatusMessage = "Moving to Correspondence folder...";
                            var (store, folder) = OutlookService.ParseOutlookPath(_currentIcSetting.OutlookInboxFolderPath);
                            new OutlookService().MoveEmailToFolder(outlookApp, _currentEmail.Emailid, $"\\\\{store}\\{folder}", _currentIcSetting.OutlookCorrespondenceFolderPath);
                            shouldAutoAdvance = true;
                            return;

                        case NoGisFilesWindow.UserChoice.Fail:
                            var noGisTest = namedTests.returnNewTestResult("GIS_No_GIS_Attachments", _currentEmail.Emailid, IcTestResult.TestType.Deliverable);
                            noGisTest.Passed = false;
                            _currentEmailTestResult.AddSubordinateTestResult(noGisTest);
                            // Fall through to the standard failure handling logic below.
                            break;

                        case NoGisFilesWindow.UserChoice.Cancel:
                        default:
                            // If the user closes the window, treat it as a skip.
                            StatusMessage = "Operation canceled by user.";
                            shouldAutoAdvance = true; // Mark to advance, but don't move the email.
                            return;
                    }
                }

                // 1. Populate ALL UI grids and lists first, regardless of the outcome.
                if (processingResult.AttachmentAnalysis?.IdentifiedFileSets?.Any() == true)
                {
                    await RunOnUIThread(() =>
                    {
                        foreach (var fs in processingResult.AttachmentAnalysis.IdentifiedFileSets)
                        {
                            var fsVM = new FileSetViewModel(fs)
                            {
                                UseFilter = !fs.filesetType.Equals("shapefile", StringComparison.OrdinalIgnoreCase)
                            };
                            _foundFileSets.Add(fsVM);
                        }
                    });
                }

                if (processingResult.ShapeItems?.Any() == true)
                {
                    _allProcessedShapes = processingResult.ShapeItems;
                }
                UpdateFileSetCounts();
                await RefreshShapeListsAndMap();
                await ZoomToAllAndSiteAsync();

                // 2. Now, check the final result and decide the next step.
                if (_currentEmailTestResult.CumulativeAction.ResultAction == TestActionResponse.Pass)
                {
                    StatusMessage = "Ready for review.";
                    IsEmailActionEnabled = true;
                }
                else
                {
                    // For any other result, show the results window and enable the action buttons.
                    ShowTestResultWindow(_currentEmailTestResult);
                    StatusMessage = $"Review required: {_currentEmailTestResult.Comments.FirstOrDefault()}";
                    IsEmailActionEnabled = true;
                }
                // --- END OF MODIFIED LOGIC ---
            }
            catch (Exception ex)
            {
                shouldAutoAdvance = true;
                Log.RecordError($"An unexpected error occurred while processing email ID {currentEmailSummary.Emailid}", ex, "ProcessSelectedQueueAsync");
                ArcGIS.Desktop.Framework.Dialogs.MessageBox.Show("An unexpected error occurred. The application will advance to the next email.", "Processing Error");
            }
            finally
            {
                if (emailToProcess != null)
                {
                    _pathForNextCleanup = emailToProcess.TempFolderPath;
                }
                if (SelectedIcType != null)
                {
                    SelectedIcType.EmailCount = emailsToProcess.Count;
                }

                if (outlookApp != null)
                {
                    Marshal.ReleaseComObject(outlookApp);
                }

                if (shouldAutoAdvance)
                {
                    SelectedIcType.FailedCount++;
                    if (emailsToProcess.Any() && emailsToProcess.First() == currentEmailSummary)
                    {
                        emailsToProcess.RemoveAt(0);
                    }
                    await ProcessNextEmail();
                }
            }
        }

        private async Task ProcessSelectedQueueAsync_bak()
        {
            // --- 1. Initial UI and Configuration Setup ---
            await PerformCleanupAsync();
            await ClearManuallyLoadedLayersAsync();
            IsEmailActionEnabled = false;
            _foundFileSets.Clear();
            _allProcessedShapes.Clear();

            if (SelectedIcType == null || !_emailQueues.TryGetValue(SelectedIcType.Name, out var emailsToProcess) || !emailsToProcess.Any())
            {
                CurrentEmailSubject = "Queue is empty.";
                StatusMessage = $"Queue '{SelectedIcType?.Name}' is empty.";
                return;
            }

            Outlook.Application outlookApp = null;
            var namedTests = new IcNamedTests(Log, PostGreTool);
            var currentEmailSummary = emailsToProcess.First();
            EmailItem emailToProcess = null;
           
            // This flag is the master controller for advancing the queue.
            bool shouldAutoAdvance = false;

            try
            {
                outlookApp = new Outlook.Application();                
                var (storeName, folderPath) = OutlookService.ParseOutlookPath(_currentIcSetting.OutlookInboxFolderPath);

                emailToProcess = await QueuedTask.Run(() => new OutlookService().GetEmailById(outlookApp, folderPath, currentEmailSummary.Emailid, storeName));
                _currentEmail = emailToProcess;

                // --- START OF NEW DUPLICATE FILENAME CHECK ---
                //if (emailToProcess != null && emailToProcess.Attachments.Any())
                //{
                //    // Find any original filenames that appear more than once.
                //    var duplicateOriginalFilenames = emailToProcess.Attachments
                //                               .GroupBy(a => a.OriginalFileName, StringComparer.OrdinalIgnoreCase)
                //                               .Where(g => g.Count() > 1)
                //                               .Select(g => g.Key)
                //                               .ToList();

                //    if (duplicateOriginalFilenames.Any())
                //    {
                //        var multiFileDuplicates = new List<string>();

                //        // For each duplicate, check if it belongs to a multi-file dataset.
                //        foreach (var dupName in duplicateOriginalFilenames)
                //        {
                //            var rule = IcRules.ReturnFilesetRuleForExtension(Path.GetExtension(dupName).TrimStart('.'));
                //            if (rule != null && rule.RequiredExtensions.Count > 1)
                //            {
                //                multiFileDuplicates.Add(dupName);
                //            }
                //        }

                //        if (multiFileDuplicates.Any())
                //        {
                //            _currentEmailTestResult = namedTests.returnNewTestResult("GIS_Root_Email_Load", emailToProcess.Emailid, IcTestResult.TestType.Deliverable);

                //            var duplicateTest = namedTests.returnNewTestResult("GIS_DuplicateFilenamesInAttachments", emailToProcess.Emailid, IcTestResult.TestType.Deliverable);
                //            duplicateTest.Passed = false;
                //            duplicateTest.AddComment($"The submission could not be processed because it contains multiple multi-file datasets with the same filename(s): {string.Join(", ", multiFileDuplicates.Distinct())}");

                //            _currentEmailTestResult.AddSubordinateTestResult(duplicateTest);

                //            ShowTestResultWindow(_currentEmailTestResult);
                //            StatusMessage = "Processing failed: Duplicate filenames found in attachments.";
                //            IsEmailActionEnabled = true; // Allow user to Reject/Skip
                //            UpdateEmailInfo(emailToProcess, new EmailClassificationResult(), false, EmailType.Unknown);
                //            return; // Stop processing this email
                //        }
                //    }
                //}
                // --- END OF NEW DUPLICATE FILENAME CHECK ---

                if (emailToProcess == null)
                {
                    shouldAutoAdvance = true; // Mark for advancement
                    ArcGIS.Desktop.Framework.Dialogs.MessageBox.Show($"Could not retrieve the email: '{currentEmailSummary.Subject}'. It will be skipped.", "Email Retrieval Error");
                    return; // Exit the try block; the finally block will handle the rest.
                }

                var classification = new EmailClassifierService(IcRules, Log).ClassifyEmail(emailToProcess);
                _currentClassification = classification;

                EmailType finalEmailType = classification.Type;
                if (classification.Type == EmailType.Unknown || classification.Type == EmailType.EmptySubjectline)
                {
                    var (wasSelected, selectedType) = await RequestManualEmailClassification(emailToProcess);
                    if (wasSelected)
                    {
                        finalEmailType = selectedType;
                        classification.WasManuallyClassified = true;
                    }
                    else
                    {
                        shouldAutoAdvance = true; // User canceled
                        return; // Exit try block
                    }
                }

                UpdateEmailInfo(emailToProcess, classification, classification.WasManuallyClassified, finalEmailType);

                _currentSiteLocation = await GetSiteCoordinatesFromNjemsAsync(CurrentPrefId);

                var processingService = new EmailProcessingService(IcRules, namedTests, Log);
                EmailProcessingResult processingResult = await processingService.ProcessEmailAsync(outlookApp, emailToProcess, classification, SelectedIcType.Name, folderPath, storeName, classification.WasManuallyClassified, finalEmailType, GetSiteCoordinatesFromPostgreAsync);

                _currentEmailTestResult = processingResult.TestResult;
                _currentAttachmentAnalysis = processingResult.AttachmentAnalysis;
               // _currentFilesetTestResults = processingResult.FilesetTestResults;

                if (processingResult.TestResult == null)
                {
                    shouldAutoAdvance = true;
                    return;
                }

                // If the result did not pass validation, show the results window and enable the
                // action buttons so the user can make a manual decision.
                if (_currentEmailTestResult.CumulativeAction.ResultAction != TestActionResponse.Pass)
                {
                    ShowTestResultWindow(_currentEmailTestResult);
                    StatusMessage = $"Review required: {_currentEmailTestResult.Comments.FirstOrDefault()}";
                    IsEmailActionEnabled = true; // Enable Save/Skip/Reject buttons
                    return; // Stop processing and wait for the user to click a button
                }

                // If we reach here, the result was a clean Pass, so we load the UI for review.
                if (processingResult.ShapeItems?.Any() == true)
                {
                    _allProcessedShapes = processingResult.ShapeItems;
                }

                if (processingResult.AttachmentAnalysis?.IdentifiedFileSets?.Any() == true)
                {
                    await RunOnUIThread(() =>
                    {
                        foreach (var fs in processingResult.AttachmentAnalysis.IdentifiedFileSets)
                        {
                            var fsVM = new FileSetViewModel(fs)
                            {
                                UseFilter = !fs.filesetType.Equals("shapefile", StringComparison.OrdinalIgnoreCase)
                            };
                            _foundFileSets.Add(fsVM);
                        }
                    });
                }

                UpdateFileSetCounts();
                await RefreshShapeListsAndMap();
                await ZoomToAllAndSiteAsync();

                StatusMessage = "Ready for review.";
                IsEmailActionEnabled = true;



                //switch (_currentEmailTestResult.CumulativeAction.ResultAction)
                //{
                //    case TestActionResponse.Pass:
                //        // 1. Store all shapes in our new master list.
                //        if (processingResult.ShapeItems?.Any() == true)
                //        {
                //            _allProcessedShapes = processingResult.ShapeItems;
                //        }

                //        // 2. Set up the FileSetViewModels and their default values.
                //        if (processingResult.AttachmentAnalysis?.IdentifiedFileSets?.Any() == true)
                //        {
                //            await RunOnUIThread(() =>
                //            {
                //                foreach (var fs in processingResult.AttachmentAnalysis.IdentifiedFileSets)
                //                {
                //                    var fsVM = new FileSetViewModel(fs);
                //                    // Set default "UseFilter" state: true for DWG, false for shapefiles.
                //                    fsVM.UseFilter = !fs.filesetType.Equals("shapefile", StringComparison.OrdinalIgnoreCase);
                //                    _foundFileSets.Add(fsVM);
                //                }
                //            });
                //        }

                //        // 3. Calculate and populate the counts for each fileset.
                //        UpdateFileSetCounts();

                //        //var shapesByFile = _allProcessedShapes.GroupBy(s => s.SourceFile);
                //        //foreach (var group in shapesByFile)
                //        //{
                //        //    var fileSetVM = _foundFileSets.FirstOrDefault(fs => fs.FileName == group.Key);
                //        //    if (fileSetVM != null)
                //        //    {
                //        //        fileSetVM.TotalFeatureCount = group.Count();
                //        //        fileSetVM.FilteredCount = group.Count(s => s.IsAutoSelected);
                //        //        fileSetVM.ValidFeatureCount = group.Count(s => s.IsValid);
                //        //        fileSetVM.InvalidFeatureCount = group.Count(s => !s.IsValid);
                //        //    }
                //        //}

                //        // 4. Call our new central refresh method. This single call now handles
                //        //    populating the UI lists and redrawing the map based on the checkbox states.
                //        await RefreshShapeListsAndMap();
                //        await ZoomToAllAndSiteAsync();

                //        // 5. Set the final UI state.
                //        StatusMessage = "Ready for review.";
                //        IsEmailActionEnabled = true;
                //        break;

                //    case TestActionResponse.Note:
                //    case TestActionResponse.Manual:
                //    case TestActionResponse.Fail:
                //    default:
                //        // Any non-passing result will auto-advance to the next email
                //        shouldAutoAdvance = true;
                //        UpdateQueueStats(_currentEmailTestResult); // Update stats based on failure type
                //        ShowTestResultWindow(_currentEmailTestResult);
                //        return;
                //}
            }
            catch (Exception ex)
            {
                shouldAutoAdvance = true; // Also advance on unexpected errors
                Log.RecordError($"An unexpected error occurred while processing email ID {currentEmailSummary.Emailid}", ex, "ProcessSelectedQueueAsync");
                ArcGIS.Desktop.Framework.Dialogs.MessageBox.Show("An unexpected error occurred. The application will advance to the next email.", "Processing Error");
            }
            finally
            {
                //CleanupTempFolder(emailToProcess);
                if (emailToProcess != null)
                {
                    _pathForNextCleanup = emailToProcess.TempFolderPath;
                }
                if (SelectedIcType != null)
                {
                    SelectedIcType.EmailCount = emailsToProcess.Count;
                }

                if (outlookApp != null)
                {
                    Marshal.ReleaseComObject(outlookApp);
                }

                // The application only advances if it was explicitly marked for auto-advancement.
                if (shouldAutoAdvance)
                {
                    SelectedIcType.FailedCount++;
                    if (emailsToProcess.Any() && emailsToProcess.First() == currentEmailSummary)
                    {
                        emailsToProcess.RemoveAt(0);
                    }
                    await ProcessNextEmail();
                }
            }
        }

        /// <summary>
        /// A helper method that simply calls the main processing logic.
        /// This will be triggered by the user action buttons.
        /// </summary>
        private async Task ProcessNextEmail()
        {
            Log.AddBlankLine();
            Log.RecordMessage(" -------------------------------------------", BisLogMessageType.Note);
            await ProcessSelectedQueueAsync();
        }

        private async Task<(bool wasSelected, EmailType selectedType)> RequestManualEmailClassification(EmailItem email)
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

            if (wasManuallySelected)
            {
                StatusMessage = $"Processing email as type '{finalType}'...";
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
                case TestActionResponse.Pass:
                    SelectedIcType.PassedCount++;
                    StatusMessage = "Email processed successfully. Ready for review.";
                    break;

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
        /// Clears all existing IC graphics and redraws the shapes from both the 'Review'
        /// and 'Use' lists with the correct symbology loaded from the project style file.
        /// </summary>
        private async Task RedrawAllShapesOnMapAsync()
        {
            // 1. Ask the SymbolManager for the symbols we need.
            var reviewSymbol = await SymbolManager.GetSymbolAsync<CIMPolygonSymbol>("ReviewShapeSymbol");
            var useSymbol = await SymbolManager.GetSymbolAsync<CIMPolygonSymbol>("UseShapeSymbol");
            var siteSymbol = await SymbolManager.GetSymbolAsync<CIMPointSymbol>("SiteLocationSymbol");

            if (reviewSymbol == null || useSymbol == null || siteSymbol == null)
            {
                // The SymbolManager will have already logged the specific error.
                return;
            }

            // Create safe copies of the lists to pass to the background thread
            List<ShapeItem> reviewShapesCopy;
            List<ShapeItem> selectedShapesCopy;
            lock (_lock)
            {
                reviewShapesCopy = _shapesToReview.ToList();
                selectedShapesCopy = _selectedShapes.ToList();
            }
            // 2. The rest of the method uses QueuedTask to draw the graphics.
            await QueuedTask.Run(() =>
            {
                var mapView = MapView.Active;
                if (mapView == null) return;

                var graphicsLayer = mapView.Map.FindLayers("IC Loader Shapes").FirstOrDefault() as GraphicsLayer;
                if (graphicsLayer == null) return;

                graphicsLayer.RemoveElements();

                // Use the safe copies for drawing
                foreach (var shapeItem in reviewShapesCopy)
                {
                    if (shapeItem.Geometry != null && !shapeItem.IsHidden)
                    {
                        graphicsLayer.AddElement(shapeItem.Geometry, reviewSymbol, shapeItem.ShapeReferenceId.ToString());
                    }
                }

                foreach (var shapeItem in selectedShapesCopy)
                {
                    if (shapeItem.Geometry != null && !shapeItem.IsHidden)
                    {
                        graphicsLayer.AddElement(shapeItem.Geometry, useSymbol, shapeItem.ShapeReferenceId.ToString());
                    }
                }

                if (_currentSiteLocation != null)
                {
                    graphicsLayer.AddElement(_currentSiteLocation, siteSymbol);
                }
                graphicsLayer.ClearSelection();
            });
        }

        /// <summary>
        /// (SHELL METHOD) Queries the database to get the coordinates for a given Preference ID.
        /// </summary>
        /// <param name="prefId">The Preference ID to search for.</param>
        /// <returns>A MapPoint object representing the site's location, or null if not found.</returns>
        private async Task<MapPoint> GetSiteCoordinatesFromPostgreAsync(string prefId)
        {
            if (string.IsNullOrWhiteSpace(prefId) || prefId.Equals("N/A", StringComparison.OrdinalIgnoreCase))
            {
                return null;
            }

            Log.RecordMessage($"Querying coordinates for Pref ID: {prefId}", BisLogMessageType.Note);
            MapPoint siteLocation = null;

            await QueuedTask.Run(() =>
            {
                var queryResult = (DataTable) Module1.PostGreTool.ExecuteNamedQuery("returnRecordedPrefId", new Dictionary<string, object> { { "@PrefID", prefId } });

                // 2. Parse the X and Y coordinates from the query result.
                if (queryResult != null && queryResult.Rows.Count > 0)
                {
                    double x = Convert.ToDouble(queryResult.Rows[0]["x_coord_spf"]);
                    double y = Convert.ToDouble(queryResult.Rows[0]["y_coord_spf"]);
                    var sr = SpatialReferenceBuilder.CreateSpatialReference(_currentIcSetting.GeometryRules.ProjectionId);
                    siteLocation = MapPointBuilder.CreateMapPoint(x, y, sr);
                }
            });

            return siteLocation;
        }
        private async Task<MapPoint> GetSiteCoordinatesFromNjemsAsync(string prefId)
        {
            if (string.IsNullOrWhiteSpace(prefId) || prefId.Equals("N/A", StringComparison.OrdinalIgnoreCase))
            {
                return null;
            }

            Log.RecordMessage($"Querying coordinates for Pref ID: {prefId} from NJEMS.", BisLogMessageType.Note);
            MapPoint siteLocation = null;
            DataTable resultTable = null;

            await QueuedTask.Run(() =>
            {
                var paramDict = new Dictionary<string, object> { { "PrefID", prefId } };
                resultTable = Module1.NjemsTool.ExecuteNamedQuery("ReturnPrefIdCoords", paramDict) as DataTable;
            });

            if (resultTable == null || resultTable.Rows.Count == 0)
            {
                Log.RecordMessage($"No coordinates found in NJEMS for Pref ID: {prefId}", BisLogMessageType.Note);
                return null;
            }

            double x = 0;
            double y = 0;

            if (resultTable.Rows.Count > 1)
            {
                string warningMsg = $"Multiple coordinates found in CORE_PI_COORDINATE_DETAIL for {prefId}. Using the first valid set.";
                Log.RecordMessage(warningMsg, BisLogMessageType.Warning);
                // Show a popup to the user on the UI thread
                FrameworkApplication.Current.Dispatcher.Invoke(() =>
                {
                    ArcGIS.Desktop.Framework.Dialogs.MessageBox.Show(warningMsg, "Multiple Coordinates Found");
                });

                // Find the first row with non-zero coordinates
                foreach (DataRow row in resultTable.Rows)
                {
                    double.TryParse(row["X_COORDINATE"]?.ToString(), out double currentX);
                    double.TryParse(row["Y_COORDINATE"]?.ToString(), out double currentY);
                    if (currentX != 0 && currentY != 0)
                    {
                        x = currentX;
                        y = currentY;
                        break; // Use the first valid set and stop looking
                    }
                }
            }
            else
            {
                // Only one row was returned
                double.TryParse(resultTable.Rows[0]["X_COORDINATE"]?.ToString(), out x);
                double.TryParse(resultTable.Rows[0]["Y_COORDINATE"]?.ToString(), out y);
            }

            // Finally, create the MapPoint if the coordinates are valid
            if (x != 0 && y != 0)
            {
                await QueuedTask.Run(() =>
                {
                    var sr = SpatialReferenceBuilder.CreateSpatialReference(_currentIcSetting.GeometryRules.ProjectionId);
                    siteLocation = MapPointBuilder.CreateMapPoint(x, y, sr);
                });

                var coordinateService = new Services.CoordinateService();
                await coordinateService.UpdatePrefIdCoordinatesInPostgresAsync(prefId, x, y, "NJEMS");
            }

            return siteLocation;
        }


    }
}