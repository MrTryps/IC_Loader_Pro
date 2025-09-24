using ArcGIS.Core.CIM;
using ArcGIS.Core.Data;
using ArcGIS.Core.Geometry;
using ArcGIS.Desktop.Core;
using ArcGIS.Desktop.Core.Geoprocessing;
using ArcGIS.Desktop.Editing;
using ArcGIS.Desktop.Framework;
using ArcGIS.Desktop.Framework.Contracts;
using ArcGIS.Desktop.Framework.Threading.Tasks;
using ArcGIS.Desktop.Internal.Mapping.Ribbon;
using ArcGIS.Desktop.Mapping;
using ArcGIS.Desktop.Mapping.Events;
using BIS_Tools_DataModels_2025;
using IC_Loader_Pro.Models;
using IC_Loader_Pro.Services;
using IC_Loader_Pro.ViewModels;
using IC_Rules_2025;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;
using System.Windows.Input;
using static BIS_Log;
using static IC_Loader_Pro.Module1;
using static IC_Rules_2025.IcTestResult;
using Outlook = Microsoft.Office.Interop.Outlook;

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
        public ICommand ShowResultsCommand { get; private set; }
        public ICommand AddSelectedShapeCommand { get; private set; }
        public ICommand RemoveSelectedShapeCommand { get; private set; }
        public ICommand AddAllShapesCommand { get; private set; }
        public ICommand RemoveAllShapesCommand { get; private set; }
        public ICommand ClearSelectionCommand { get; private set; }
        public ICommand ZoomToAllCommand { get; private set; }
        public ICommand ZoomToSelectedReviewShapeCommand { get; private set; }
        public ICommand ZoomToSelectedUseShapeCommand { get; private set; }
        public ICommand ZoomToSiteCommand { get; private set; }
        public ICommand ActivateSelectToolCommand { get; private set; }
        public ICommand HideSelectionCommand { get; private set; }
        public ICommand UnhideAllCommand { get; private set; }
        public ICommand LoadFileSetCommand { get; private set; }
        public ICommand ReloadFileSetCommand { get; private set; }
        public ICommand AddSubmissionCommand { get; private set; }
        public ICommand CreateNewIcDeliverableCommand { get; private set; }
        public ICommand OpenConnectionTesterCommand { get; private set; }
        public ICommand OpenEmailInOutlookCommand { get; private set; }
        public ICommand ProcessManualLayerCommand { get; private set; }

        #endregion

        #region Command Methods
        // this version has issues writting to shapeinfo... working to make the reast of the code work first
        //private async Task OnSave()    
        //{
        //    if (_currentEmailTestResult == null || !_selectedShapes.Any())
        //    {
        //        ArcGIS.Desktop.Framework.Dialogs.MessageBox.Show(
        //            "There must be at least one shape in the 'Selected Shapes to Use' list to save.",
        //            "Save Error");
        //        return;
        //    }

        //    IsEmailActionEnabled = false;
        //    StatusMessage = "Saving... Please wait.";
        //    Log.RecordMessage("Save process started.", BisLogMessageType.Note);

        //    var deliverableService = new Services.DeliverableService();
        //    var submissionService = new Services.SubmissionService();
        //    var shapeService = new Services.ShapeProcessingService(IcRules, Log);
        //    var notificationService = new Services.NotificationService();
        //    var outlookService = new Services.OutlookService();
        //    var testResultService = new Services.TestResultService();
        //    Outlook.Application outlookApp = null; // Declared here so it's accessible in the 'finally' block

        //    string newDelId = null;
        //    try
        //    {
        //        // === Step 1: Create Deliverable Record ===
        //        var goodCounts = new Dictionary<string, int>();
        //        var dupCounts = new Dictionary<string, int>();
        //        StatusMessage = "Creating deliverable record...";
        //        newDelId = await deliverableService.CreateNewDeliverableRecordAsync(
        //            "EMAIL", SelectedIcType.Name, CurrentPrefId, _currentEmail.ReceivedTime);
        //        CurrentDelId = newDelId;

        //        // === Step 2: Save Email, Contact, and Body Data ===
        //        await deliverableService.UpdateEmailInfoRecordAsync(newDelId, _currentEmail, _currentClassification, _currentIcSetting.OutlookInboxFolderPath);
        //        await deliverableService.UpdateContactInfoRecordAsync(newDelId, _currentEmail);
        //        var bodyParser = new Services.EmailBodyParserService(SelectedIcType.Name);
        //        var bodyData = bodyParser.GetFieldsFromBody(_currentEmail.Body);
        //        await deliverableService.UpdateBodyDataRecordAsync(newDelId, bodyData);
        //        StatusMessage = $"Successfully created Deliverable ID: {newDelId}";
        //        Log.RecordMessage(StatusMessage, BisLogMessageType.Note);
        //        if (Module1.IsInTestMode)
        //        {
        //            var notesService = new Services.NotesService();
        //            await notesService.RecordNoteAsync(newDelId, "This is a test deliverable created in Test Mode.");
        //            Log.RecordMessage($"Recorded 'Test Mode' note for deliverable {newDelId}.", BisLogMessageType.Note);
        //        }


        //        // === Step 3: Record Submissions (filesets) to get their IDs ===
        //        StatusMessage = "Recording submissions...";
        //        var submissionIdMap = await submissionService.RecordSubmissionsAsync(
        //            newDelId, SelectedIcType.Name, _currentAttachmentAnalysis.IdentifiedFileSets);

        //        // === Step 4: Record all individual physical files ===
        //        await submissionService.RecordPhysicalFilesAsync(newDelId, _currentAttachmentAnalysis.AllFiles, submissionIdMap);

        //        // === Step 5: Process and record each approved shape ===
        //        foreach (var shapeToSave in _selectedShapes)
        //        {
        //            StatusMessage = $"Processing shape {shapeToSave.ShapeReferenceId}...";
        //            string submissionId = submissionIdMap.GetValueOrDefault(shapeToSave.SourceFile);
        //            if (string.IsNullOrEmpty(submissionId))
        //            {
        //                Log.RecordError($"Could not find a submission ID for shape from file '{shapeToSave.SourceFile}'. Skipping shape record.", null, nameof(OnSave));
        //                continue;
        //            }

        //            string newShapeId = await shapeService.GetNextShapeIdAsync(newDelId, _currentIcSetting.IdPrefix);

        //            bool recordCreated = await shapeService.RecordShapeInfoAsync(newShapeId, submissionId, newDelId, CurrentPrefId, SelectedIcType.Name);
        //            if (!recordCreated)
        //            {
        //                Log.RecordError($"Aborting processing for shape from file '{shapeToSave.SourceFile}' because its info record could not be created.", null, nameof(OnSave));
        //                continue; // Move to the next shape
        //            }

        //            bool isDuplicate = await shapeService.IsDuplicateInProposedAsync(shapeToSave.Geometry, CurrentPrefId, SelectedIcType.Name);

        //            if (isDuplicate)
        //            {
        //                if (!dupCounts.ContainsKey(submissionId)) dupCounts[submissionId] = 0;
        //                dupCounts[submissionId]++;
        //                await shapeService.UpdateShapeInfoFieldAsync(newShapeId, "SHAPE_STATUS", "Duplicate", SelectedIcType.Name);
        //                Log.RecordMessage($"Shape {newShapeId} was found to be a duplicate.", BisLogMessageType.Note);
        //            }
        //            else
        //            {
        //                if (!goodCounts.ContainsKey(submissionId)) goodCounts[submissionId] = 0;
        //                goodCounts[submissionId]++;
        //                await shapeService.UpdateShapeInfoFieldAsync(newShapeId, "SHAPE_STATUS", "To Be Reviewed", SelectedIcType.Name);
        //            }

        //            // Record additional shape metadata
        //            await shapeService.UpdateShapeInfoFieldAsync(newShapeId, "CREATED_BY", "Crawler", SelectedIcType.Name);
        //            await shapeService.UpdateShapeInfoFieldAsync(newShapeId, "CENTROID_X", shapeToSave.Geometry.Extent.Center.X, SelectedIcType.Name);
        //            await shapeService.UpdateShapeInfoFieldAsync(newShapeId, "CENTROID_Y", shapeToSave.Geometry.Extent.Center.Y, SelectedIcType.Name);
        //            await shapeService.UpdateShapeInfoFieldAsync(newShapeId, "SITE_DIST", shapeToSave.DistanceFromSite, SelectedIcType.Name);

        //            // TODO: Logic for recording notes/comments for the individual shape can be added here.

        //            // Copy the shape's geometry into the 'proposed' feature class
        //            await shapeService.CopyShapeToProposedAsync(shapeToSave.Geometry, newShapeId, SelectedIcType.Name);
        //        }

        //        // === Step 6: Move Files to Final Location ===
        //        StatusMessage = "Archiving submission files...";
        //        await submissionService.MoveAllSubmissionsAsync(
        //            _currentAttachmentAnalysis.IdentifiedFileSets, submissionIdMap, _currentIcSetting.AsSubmittedPath);

        //        // === Step 7: Finalize, Notify, and Clean Up ===
        //        // ... (Update counts, deliverable status, send email, move source email...)
        //        foreach (var subId in submissionIdMap.Values)
        //        {
        //            await submissionService.UpdateSubmissionCountsAsync(subId, goodCounts.GetValueOrDefault(subId, 0), dupCounts.GetValueOrDefault(subId, 0));
        //        }

        //        string finalStatus = (goodCounts.Values.Sum() == 0 && dupCounts.Values.Sum() > 0) ? "Duplicate" : "Migrated";
        //        await deliverableService.UpdateDeliverableStatusAsync(newDelId, finalStatus, "Pass");
        //        await testResultService.SaveTestResultsAsync(_currentEmailTestResult, newDelId);
        //        await notificationService.SendConfirmationEmailAsync(newDelId);

        //        StatusMessage = "Moving processed email...";
        //        outlookApp = new Outlook.Application();
        //        var (storeName, folderPath) = OutlookService.ParseOutlookPath(_currentIcSetting.OutlookInboxFolderPath);
        //        outlookService.MoveEmailToFolder(outlookApp, _currentEmail.Emailid, folderPath, storeName, _currentIcSetting.OutlookProcessedFolderPath);

        //        StatusMessage = $"Successfully saved submission as {newDelId}.";
        //        SelectedIcType.PassedCount++;

        //    }
        //    catch (Exception ex)
        //    {
        //        Log.RecordError("An error occurred while creating the new deliverable record.", ex, nameof(OnSave));
        //        ArcGIS.Desktop.Framework.Dialogs.MessageBox.Show(
        //            "An error occurred while saving the submission. Please check the logs.",
        //            "Save Error");
        //        StatusMessage = "Save failed. Please review logs.";
        //        IsEmailActionEnabled = true; // Re-enable buttons on failure
        //        return; // Stop the save process
        //    }

        //    // If the save was successful, update the stats and advance to the next email.
        //    SelectedIcType.PassedCount++;

        //    if (_emailQueues.TryGetValue(SelectedIcType.Name, out var emailsToProcess) && emailsToProcess.Any())
        //    {
        //        emailsToProcess.RemoveAt(0);
        //    }
        //    await ProcessNextEmail();
        //}

        //private async Task OnSave()
        //{
        //    if (_currentEmailTestResult == null || !_selectedShapes.Any())
        //    {
        //        ArcGIS.Desktop.Framework.Dialogs.MessageBox.Show(
        //            "There must be at least one shape in the 'Selected Shapes to Use' list to save.", "Save Error");
        //        return;
        //    }

        //    IsEmailActionEnabled = false;
        //    StatusMessage = "Saving... Please wait.";
        //    Log.RecordMessage("Save process started.", BisLogMessageType.Note);

        //    var deliverableService = new Services.DeliverableService();
        //    var submissionService = new Services.SubmissionService();
        //    var shapeService = new Services.ShapeProcessingService(IcRules, Log);
        //    var notificationService = new Services.NotificationService();
        //    var outlookService = new Services.OutlookService();
        //    var testResultService = new Services.TestResultService();
        //    Outlook.Application outlookApp = null;

        //    string newDelId = null;
        //    try
        //    {
        //        // === Step 1 & 2: Create Deliverable and Save Metadata ===
        //        outlookApp = new Outlook.Application();
        //        var goodCounts = new Dictionary<string, int>();
        //        var dupCounts = new Dictionary<string, int>();
        //        StatusMessage = "Creating deliverable record...";
        //        newDelId = await deliverableService.CreateNewDeliverableRecordAsync(
        //            "EMAIL", SelectedIcType.Name, CurrentPrefId, _currentEmail.ReceivedTime);
        //        CurrentDelId = newDelId;
        //        _currentEmailTestResult.RefId = newDelId;
        //        _currentEmailTestResult.addParameter("prefid", CurrentPrefId);

        //        await deliverableService.UpdateEmailInfoRecordAsync(newDelId, _currentEmail, _currentClassification, _currentIcSetting.OutlookInboxFolderPath);
        //        await deliverableService.UpdateContactInfoRecordAsync(newDelId, _currentEmail);
        //        var bodyParser = new Services.EmailBodyParserService(SelectedIcType.Name);
        //        var bodyData = bodyParser.GetFieldsFromBody(_currentEmail.Body);
        //        await deliverableService.UpdateBodyDataRecordAsync(newDelId, bodyData);

        //        // === Step 3: Record Submissions (filesets) ===
        //        StatusMessage = "Recording submissions...";
        //        var submissionIdMap = await submissionService.RecordSubmissionsAsync(
        //            newDelId, SelectedIcType.Name, _currentAttachmentAnalysis.IdentifiedFileSets);
        //        await submissionService.RecordPhysicalFilesAsync(newDelId, _currentAttachmentAnalysis.AllFiles, submissionIdMap);

        //        // === Step 4: Copy Approved Shapes to the 'Proposed' Feature Class ===

        //        foreach (var shapeToSave in _selectedShapes)
        //        {
        //            StatusMessage = $"Processing shape {shapeToSave.ShapeReferenceId}...";
        //            string submissionId = submissionIdMap.GetValueOrDefault(shapeToSave.SourceFile);
        //            if (string.IsNullOrEmpty(submissionId))
        //            {
        //                Log.RecordError($"Could not find a submission ID for shape from file '{shapeToSave.SourceFile}'. Skipping shape record.", null, nameof(OnSave));
        //                continue;
        //            }

        //            // Get the next unique Shape ID from the database service.
        //            string newShapeId = await shapeService.GetNextShapeIdAsync(newDelId, _currentIcSetting.IdPrefix);
        //            // --- END OF CORRECTION ---

        //            bool recordCreated = await shapeService.RecordShapeInfoAsync(newShapeId, submissionId, newDelId, CurrentPrefId, SelectedIcType.Name);
        //            if (!recordCreated)
        //            {
        //                Log.RecordError($"Aborting processing for this shape because its info record could not be created.", null, nameof(OnSave));
        //                continue; // Skip to the next shape if the info record fails
        //            }

        //            bool isDuplicate = await shapeService.IsDuplicateInProposedAsync(shapeToSave.Geometry, CurrentPrefId, SelectedIcType.Name);

        //            if (isDuplicate)
        //            {
        //                if (!dupCounts.ContainsKey(submissionId)) dupCounts[submissionId] = 0;
        //                dupCounts[submissionId]++;
        //                await shapeService.UpdateShapeInfoFieldAsync(newShapeId, "SHAPE_STATUS", "Duplicate", SelectedIcType.Name);
        //            }
        //            else
        //            {
        //                if (!goodCounts.ContainsKey(submissionId)) goodCounts[submissionId] = 0;
        //                goodCounts[submissionId]++;
        //                await shapeService.UpdateShapeInfoFieldAsync(newShapeId, "SHAPE_STATUS", "To Be Reviewed", SelectedIcType.Name);
        //            }

        //            await shapeService.UpdateShapeInfoFieldAsync(newShapeId, "CREATED_BY", "Crawler", SelectedIcType.Name);
        //            await shapeService.UpdateShapeInfoFieldAsync(newShapeId, "CENTROID_X", shapeToSave.Geometry.Extent.Center.X, SelectedIcType.Name);
        //            await shapeService.UpdateShapeInfoFieldAsync(newShapeId, "CENTROID_Y", shapeToSave.Geometry.Extent.Center.Y, SelectedIcType.Name);
        //            await shapeService.UpdateShapeInfoFieldAsync(newShapeId, "SITE_DIST", shapeToSave.DistanceFromSite, SelectedIcType.Name);

        //            await shapeService.CopyShapeToProposedAsync(shapeToSave.Geometry, newShapeId, SelectedIcType.Name);
        //        }

        //        // === Step 5: Finalize, Record Results, and Clean Up ===
        //        StatusMessage = "Finalizing records...";

        //        var finalTestResult = testResultService.CompileFinalResults(
        //    _currentEmailTestResult,
        //    _currentFilesetTestResults,
        //    _selectedShapes,
        //    newDelId,
        //    SelectedIcType.Name,
        //    CurrentPrefId);

        //        finalTestResult.UpdateAllRefIds(newDelId);
        //        finalTestResult.addParameter("prefid", CurrentPrefId);

        //        // Update the final status for the deliverable.
        //        await deliverableService.UpdateDeliverableStatusAsync(newDelId, "Migrated", "Pass");

        //        // Record the final, compiled test results to the database.
        //        await testResultService.SaveTestResultsAsync(finalTestResult, newDelId);

        //        // Send confirmation email (shelled).
        //        await notificationService.SendConfirmationEmailAsync(newDelId, finalTestResult, SelectedIcType.Name, outlookApp);

        //        // Move the processed email.
        //        StatusMessage = "Moving processed email...";
        //        outlookApp = new Outlook.Application();
        //        // The refactored method now expects the full source and destination paths.
        //        outlookService.MoveEmailToFolder(
        //            outlookApp,
        //            _currentEmail.Emailid,
        //            _currentIcSetting.OutlookInboxFolderPath,    // Full source path
        //            _currentIcSetting.OutlookProcessedFolderPath // Full destination path
        //        );

        //        StatusMessage = $"Successfully saved submission as {newDelId}.";
        //        SelectedIcType.PassedCount++;
        //    }
        //    catch (Exception ex)
        //    {
        //        Log.RecordError("A critical error occurred during the save process.", ex, nameof(OnSave));
        //        ArcGIS.Desktop.Framework.Dialogs.MessageBox.Show("An error occurred while saving. Please check the logs.", "Save Error");
        //        IsEmailActionEnabled = true;
        //        return;
        //    }
        //    finally
        //    {
        //        if (outlookApp != null)
        //        {
        //            System.Runtime.InteropServices.Marshal.ReleaseComObject(outlookApp);
        //            outlookApp = null;
        //        }
        //    }

        //    // Advance to the next email
        //    if (_emailQueues.TryGetValue(SelectedIcType.Name, out var emailsToProcess) && emailsToProcess.Any())
        //    {
        //        emailsToProcess.RemoveAt(0);
        //    }
        //    await ProcessNextEmail();
        //}


        private async Task OnSave()
        {
            if (_currentEmailTestResult == null || !_selectedShapes.Any())
            {
                ArcGIS.Desktop.Framework.Dialogs.MessageBox.Show(
                    "There must be at least one shape in the 'Selected Shapes to Use' list to save.", "Save Error");
                return;
            }
            // the new block - may need to remove
            var testResultService = new Services.TestResultService();
            var finalTestResult = testResultService.CompileFinalResults(
                _currentEmailTestResult,                
                _selectedShapes,
                "TEMP_SAVE_ID", // This will be updated by the finalizer
                SelectedIcType.Name,
                CurrentPrefId);


            // Call the generic finalizer with 'true' for an approved submission.
            await FinalizeSubmissionAsync(finalTestResult);
            SelectedIcType.PassedCount++;
        }

        private async Task OnSkip()
        {
            Log.RecordMessage("Skip button was clicked.", BisLogMessageType.Note);

            if (SelectedIcType == null || string.IsNullOrEmpty(CurrentEmailId))
            {
                StatusMessage = "Nothing to skip.";
                return;
            }

            // 1. Create a specific test result for the skip action using the new rule.
            var namedTests = new IcNamedTests(Log, PostGreTool);
            var skipTestResult = namedTests.returnNewTestResult(
                "GIS_Skipped", 
                CurrentEmailId,
                IcTestResult.TestType.Deliverable
            );
            // The test itself didn't "fail", the user just skipped it.
            skipTestResult.Passed = true;
            skipTestResult.AddComment($"User manually skipped email: {CurrentEmailSubject}");

            // (Optional) You could record this skip action to the database if needed.
            // skipTestResult.RecordResults();

            // 2. Update the queue statistics.
            // UpdateQueueStats(skipTestResult);
            SelectedIcType.SkippedCount++;

            // 3. Advance to the next email.
            if (_emailQueues.TryGetValue(SelectedIcType.Name, out var emailsToProcess) && emailsToProcess.Any())
            {
                emailsToProcess.RemoveAt(0);
            }
            await ProcessNextEmail();
        }

        //private async Task OnReject()
        //{
        //    Log.RecordMessage("Reject button was clicked.", BisLogMessageType.Note);
        //    Outlook.Application outlookApp = null;          
        //    if (SelectedIcType == null || string.IsNullOrEmpty(CurrentEmailId))
        //    {
        //        StatusMessage = "Nothing to reject.";
        //        return;
        //    }

        //    // Disable the action buttons while processing
        //    IsEmailActionEnabled = false;
        //    StatusMessage = "Processing rejection...";

        //    try
        //    {
        //        outlookApp = new Outlook.Application();
        //        // 1. Create the final test result for the manual rejection.
        //        var namedTests = new IcNamedTests(Log, PostGreTool);
        //        var rejectionTestResult = namedTests.returnNewTestResult(
        //            "GIS_Root_Email_Load",
        //            CurrentEmailId,
        //            IcTestResult.TestType.Deliverable
        //        );
        //        rejectionTestResult.Passed = false;
        //        rejectionTestResult.AddComment("Submission was manually rejected by the user.");

        //        // 2. Update the UI stats immediately.
        //        SelectedIcType.FailedCount++;

        //        // 3. Call the shared rejection logic in the service.
        //        var processingService = new EmailProcessingService(IcRules, namedTests, Log);

        //        // Get the source folder info needed to move the email.
        //        string fullOutlookPath = _currentIcSetting.OutlookInboxFolderPath;
        //        var (storeName, folderPath) = OutlookService.ParseOutlookPath(fullOutlookPath);

        //        // This call is now much cleaner and delegates the work.
        //        // It needs to run on a background thread to avoid freezing the UI.
        //        await QueuedTask.Run(() =>
        //            processingService.HandleRejection(outlookApp,rejectionTestResult, _currentIcSetting, folderPath, storeName)
        //        );
        //    }
        //    catch (Exception ex)
        //    {
        //        Log.RecordError("An error occurred during the rejection process.", ex, nameof(OnReject));
        //        ArcGIS.Desktop.Framework.Dialogs.MessageBox.Show(
        //            "An error occurred during the rejection process. Please check the logs.",
        //            "Rejection Error",
        //            System.Windows.MessageBoxButton.OK,
        //            System.Windows.MessageBoxImage.Error);
        //    }
        //    finally {                 // Ensure the Outlook application is released properly.
        //        if (outlookApp != null)
        //        {
        //            System.Runtime.InteropServices.Marshal.ReleaseComObject(outlookApp);
        //            outlookApp = null;
        //        }
        //    }

        //    // 4. Advance to the next email in the queue.
        //    if (_emailQueues.TryGetValue(SelectedIcType.Name, out var emailsToProcess) && emailsToProcess.Any())
        //    {
        //        emailsToProcess.RemoveAt(0);
        //    }
        //    await ProcessNextEmail();
        //}

        private async Task OnReject()
        {
            if (SelectedIcType == null || string.IsNullOrEmpty(CurrentEmailId))
            {
                StatusMessage = "Nothing to reject.";
                return;
            }

            // 1. Compile the complete set of test results that were generated during processing.
            var testResultService = new Services.TestResultService();
            var finalTestResult = testResultService.CompileFinalResults(
                _currentEmailTestResult,                
                _selectedShapes,
                "TEMP_REJECT_ID", // Temporary ID, will be updated by the finalizer
                SelectedIcType.Name,
                CurrentPrefId);

            // 2. Mark the final result as failed and add the specific manual rejection note.
            finalTestResult.Passed = false;
            //finalTestResult.AddComment($"Submission was manually rejected by user: {Environment.UserName}");

            // Call the generic finalizer with the complete, failed test result object.
            await FinalizeSubmissionAsync(finalTestResult);
            SelectedIcType.FailedCount++;
        }


        private async Task OnProcessManualLayerAsync()
        {
            Log.RecordMessage("Processing features from 'manually_added' layer.", BisLogMessageType.Note);
            IsUIEnabled = false;
            StatusMessage = "Processing manual features...";

            try
            {
                List<ShapeItem> manualShapes = new List<ShapeItem>();
                string sourceFileName = "manually_added";

                // This must run on the main thread
                await QueuedTask.Run(() =>
                {
                    if (_manualAddLayer == null)
                    {
                        Log.RecordError("The 'manually_added' layer could not be found in the map.", null, "OnProcessManualLayerAsync");
                        return;
                    }

                    using (var featureClass = _manualAddLayer.GetFeatureClass())
                    using (var cursor = featureClass.Search(null, false))
                    {
                        while (cursor.MoveNext())
                        {
                            using (var feature = cursor.Current as Feature)
                            {
                                if (feature?.GetShape() is Polygon polygon)
                                {
                                    var shapeItem = new ShapeItem
                                    {
                                        Geometry = polygon,
                                        SourceFile = sourceFileName,
                                        ShapeReferenceId = (int)feature.GetObjectID(),
                                        IsValid = true,
                                        Status = "Manual Add",
                                        IsAutoSelected = true // Assume manually added shapes are always intended for use
                                    };
                                    manualShapes.Add(shapeItem);
                                }
                            }
                        }
                    }
                });

                if (!manualShapes.Any())
                {
                    ArcGIS.Desktop.Framework.Dialogs.MessageBox.Show("No polygon features found in the 'manually_added' layer.", "No Features");
                    return;
                }

                // Create a new "fileset" in the UI to represent this manual submission
                var manualFileSet = new fileset
                {
                    fileName = sourceFileName,
                    filesetType = "shapefile", // Treat it as a shapefile for processing
                    validSet = true,
                    path = "manual" // A special path to identify it later
                };
                var fileSetVM = new FileSetViewModel(manualFileSet);
                _foundFileSets.Add(fileSetVM);
                _currentAttachmentAnalysis.IdentifiedFileSets.Add(manualFileSet);

                // Add the new shapes to the master list and refresh the UI
                _allProcessedShapes.AddRange(manualShapes);
                UpdateFileSetCounts();
                await RefreshShapeListsAndMap();

                StatusMessage = $"Added {manualShapes.Count} features from the manual layer.";
            }
            catch (Exception ex)
            {
                Log.RecordError("An error occurred while processing the manual layer.", ex, nameof(OnProcessManualLayerAsync));
                StatusMessage = "Error processing manual layer.";
            }
            finally
            {
                IsUIEnabled = true;
            }
        }


        private async Task OnShowNotes()
        {
            //Log.RecordMessage("Menu: Notes was clicked.", BisLogMessageType.Note);
            Log.Open();
            await Task.CompletedTask;
        }

        private async Task OnSearch()
        {
            Log.RecordMessage("Menu: Search was clicked.", BisLogMessageType.Note);
            await Task.CompletedTask;
        }

        private async Task OnTools()
        {
            Log.RecordMessage("Menu: Tools was clicked.", BisLogMessageType.Note);
            await Task.CompletedTask;
        }

        private async Task OnOptions()
        {
            Log.RecordMessage("Menu: Options was clicked.", BisLogMessageType.Note);
            await Task.CompletedTask;
        }

        private void OnShowResults()
        {
        
            ShowTestResultWindow(_currentEmailTestResult);
        }

        public static void ShowTestResultWindow(IcTestResult testResult)
        {
            if (testResult == null)
            {
                ArcGIS.Desktop.Framework.Dialogs.MessageBox.Show("No test results are available to display.", "Show Results");
                return;
            }

            var testResultViewModel = new ViewModels.TestResultViewModel(testResult);
            var testResultWindow = new Views.TestResultWindow
            {
                // We wrap the root ViewModel in a list so the TreeView can display it
                DataContext = new { RootResult = new List<ViewModels.TestResultViewModel> { testResultViewModel } },
                Owner = FrameworkApplication.Current.MainWindow
            };
            testResultWindow.ShowDialog();
        }

        protected override void OnShow(bool isVisible)
        {
 
            base.OnShow(isVisible);
        }

        #endregion

        #region Shape Manipulation Commands

        private async Task AddSelectedShape()
        {
            var itemsToMove = SelectedShapesForReview.OfType<ShapeItem>().ToList();
            if (!itemsToMove.Any()) return;

            await RunOnUIThread(() =>
            {
                foreach (var item in itemsToMove)
                {
                    _selectedShapes.Add(item);
                    _shapesToReview.Remove(item);
                }
            });
            // After moving the items, redraw the map to update their symbols.
            await RedrawAllShapesOnMapAsync();
        }

        private async Task RemoveSelectedShape()
        {
            var itemsToMove = SelectedShapesToUse.OfType<ShapeItem>().ToList();
            if (!itemsToMove.Any()) return;

            await RunOnUIThread(() =>
            {
                foreach (var item in itemsToMove)
                {
                    _shapesToReview.Add(item);
                    _selectedShapes.Remove(item);
                }
            });
            // After moving the items, redraw the map to update their symbols.
            await RedrawAllShapesOnMapAsync();
        }
             

        private async Task OnZoomToSelectedReviewShape()
        {
            // Use LINQ to get a collection of geometries from the selected items in the "Review" list.
            if (!SelectedShapesForReview.Any()) return;
            var selectedGeometries = SelectedShapesForReview.OfType<ShapeItem>().Select(s => s.Geometry);
            await ZoomToGeometryAsync(selectedGeometries);
            await RedrawAllShapesOnMapAsync();
        }

        private async Task OnZoomToSelectedUseShape()
        {
            // Use LINQ to get a collection of geometries from the selected items in the "Use" list.
            if (!SelectedShapesToUse.Any()) return;
            var selectedGeometries = SelectedShapesToUse.OfType<ShapeItem>().Select(s => s.Geometry);
            await ZoomToGeometryAsync(selectedGeometries);
            await RedrawAllShapesOnMapAsync();
        }

        private async Task OnZoomToSiteAsync()
        {
            if (_currentSiteLocation != null)
            {
                // --- START OF CORRECTED LOGIC ---
                // All map interactions must run on the main CIM thread via QueuedTask.
                await QueuedTask.Run(() =>
                {
                    var mapView = MapView.Active;
                    if (mapView == null) return;

                    // 1. Get the current camera to preserve its heading, pitch, and roll.
                    var currentCamera = mapView.Camera;

                    // 2. Create a new camera centered on our site location, with the new scale,
                    //    but using the existing camera's orientation.
                    var newCamera = new Camera(
                 _currentSiteLocation.X,
                 _currentSiteLocation.Y,
                 ZoomToSiteDistance * 2, // The new scale
                 currentCamera.Heading,  // Preserve existing heading
                 mapView.Map.SpatialReference,    // Preserve existing pitch
                 CameraViewpoint.LookAt // This is the correct type for the final argument
             );

                    // 3. Zoom to the new camera position. This is the correct method call.
                    mapView.ZoomTo(newCamera);
                });
                // --- END OF CORRECTED LOGIC ---
            }
        }

        private async Task OnOpenEmailInOutlook()
        {
            if (_currentEmail == null || _currentIcSetting == null) return;

            StatusMessage = "Opening email in Outlook...";
            IsUIEnabled = false;

            try
            {
                // Run the Outlook interaction on a background thread
                await QueuedTask.Run(() =>
                {
                    Outlook.Application outlookApp = null;
                    try
                    {
                        outlookApp = new Outlook.Application();
                        var outlookService = new OutlookService();
                        outlookService.DisplayEmailById(outlookApp, _currentEmail.Emailid, _currentIcSetting.OutlookInboxFolderPath);
                    }
                    finally
                    {
                        if (outlookApp != null) Marshal.ReleaseComObject(outlookApp);
                    }
                });
            }
            catch (Exception ex)
            {
                Log.RecordError("Failed to open email in Outlook.", ex, nameof(OnOpenEmailInOutlook));
                ArcGIS.Desktop.Framework.Dialogs.MessageBox.Show("Could not open the email in Outlook. Please ensure Outlook is running.", "Error");
            }
            finally
            {
                StatusMessage = "Ready.";
                IsUIEnabled = true;
            }
        }




        //private async Task FinalizeSubmissionAsync(bool wasApproved)
        //{
        //    IsEmailActionEnabled = false;
        //    StatusMessage = "Finalizing submission...";
        //    Log.RecordMessage("Finalization process started.", BisLogMessageType.Note);

        //    var deliverableService = new Services.DeliverableService();
        //    var submissionService = new Services.SubmissionService();
        //    var shapeService = new Services.ShapeProcessingService(IcRules, Log);
        //    var notificationService = new Services.NotificationService();
        //    var outlookService = new Services.OutlookService();
        //    var testResultService = new Services.TestResultService();
        //    Outlook.Application outlookApp = null;

        //    string newDelId = null;
        //    try
        //    {
        //        var goodCounts = new Dictionary<string, int>();
        //        var dupCounts = new Dictionary<string, int>();
        //        outlookApp = new Outlook.Application();

        //        // 1. Create the main deliverable record
        //        newDelId = await deliverableService.CreateNewDeliverableRecordAsync(
        //            "EMAIL", SelectedIcType.Name, CurrentPrefId, _currentEmail.ReceivedTime);
        //        CurrentDelId = newDelId;

        //        // 1a. Now, compile the final test result based on whether it was a pass or reject.
        //        IcTestResult finalTestResult;
        //        if (wasApproved)
        //        {
        //            finalTestResult = testResultService.CompileFinalResults(
        //                _currentEmailTestResult,
        //                _currentFilesetTestResults,
        //                _selectedShapes,
        //                newDelId,
        //                SelectedIcType.Name,
        //                CurrentPrefId);
        //        }
        //        else
        //        {
        //            finalTestResult = new IcNamedTests(Log, PostGreTool).returnNewTestResult(
        //                "GIS_Root_Email_Load",
        //                newDelId, // Use the new deliverable ID
        //                IcTestResult.TestType.Deliverable
        //            );
        //            finalTestResult.Passed = false;
        //            finalTestResult.AddComment("Submission was manually rejected by the user.");
        //        }


        //        // 2. Update the RefId for all tests to use the new permanent ID
        //        finalTestResult.UpdateAllRefIds(newDelId);

        //        // 3. Save all metadata
        //        await deliverableService.UpdateEmailInfoRecordAsync(newDelId, _currentEmail, _currentClassification, _currentIcSetting.OutlookInboxFolderPath);
        //        await deliverableService.UpdateContactInfoRecordAsync(newDelId, _currentEmail);
        //        var bodyParser = new Services.EmailBodyParserService(SelectedIcType.Name);
        //        var bodyData = bodyParser.GetFieldsFromBody(_currentEmail.Body);
        //        await deliverableService.UpdateBodyDataRecordAsync(newDelId, bodyData);

        //        // 4. Record the submission filesets to get their IDs
        //        var submissionIdMap = await submissionService.RecordSubmissionsAsync(
        //            newDelId, SelectedIcType.Name, _currentAttachmentAnalysis.IdentifiedFileSets);
        //        await submissionService.RecordPhysicalFilesAsync(newDelId, _currentAttachmentAnalysis.AllFiles, submissionIdMap);

        //        // 5. If the submission was approved, process and save the shapes
        //        int goodCount = 0;
        //        int dupCount = 0;
        //        if (wasApproved)
        //        {
        //            foreach (var shapeToSave in _selectedShapes)
        //            {
        //                StatusMessage = $"Processing shape {shapeToSave.ShapeReferenceId}...";
        //                string submissionId = submissionIdMap.GetValueOrDefault(shapeToSave.SourceFile);
        //                if (string.IsNullOrEmpty(submissionId))
        //                {
        //                    Log.RecordError($"Could not find a submission ID for shape from file '{shapeToSave.SourceFile}'. Skipping shape record.", null, nameof(OnSave));
        //                    continue;
        //                }

        //                // Get the next unique Shape ID from the database service.
        //                string newShapeId = await shapeService.GetNextShapeIdAsync(newDelId, _currentIcSetting.IdPrefix);

        //                bool recordCreated = await shapeService.RecordShapeInfoAsync(newShapeId, submissionId, newDelId, CurrentPrefId, SelectedIcType.Name);
        //                if (!recordCreated)
        //                {
        //                    Log.RecordError($"Aborting processing for this shape because its info record could not be created.", null, nameof(OnSave));
        //                    continue; // Skip to the next shape if the info record fails
        //                }

        //                bool isDuplicate = await shapeService.IsDuplicateInProposedAsync(shapeToSave.Geometry, CurrentPrefId, SelectedIcType.Name);

        //                if (isDuplicate)
        //                {
        //                    if (!dupCounts.ContainsKey(submissionId)) dupCounts[submissionId] = 0;
        //                    dupCounts[submissionId]++;
        //                    await shapeService.UpdateShapeInfoFieldAsync(newShapeId, "SHAPE_STATUS", "Duplicate", SelectedIcType.Name);
        //                }
        //                else
        //                {
        //                    if (!goodCounts.ContainsKey(submissionId)) goodCounts[submissionId] = 0;
        //                    goodCounts[submissionId]++;
        //                    await shapeService.UpdateShapeInfoFieldAsync(newShapeId, "SHAPE_STATUS", "To Be Reviewed", SelectedIcType.Name);
        //                }

        //                await shapeService.UpdateShapeInfoFieldAsync(newShapeId, "CREATED_BY", "Crawler", SelectedIcType.Name);
        //                await shapeService.UpdateShapeInfoFieldAsync(newShapeId, "CENTROID_X", shapeToSave.Geometry.Extent.Center.X, SelectedIcType.Name);
        //                await shapeService.UpdateShapeInfoFieldAsync(newShapeId, "CENTROID_Y", shapeToSave.Geometry.Extent.Center.Y, SelectedIcType.Name);
        //                await shapeService.UpdateShapeInfoFieldAsync(newShapeId, "SITE_DIST", shapeToSave.DistanceFromSite, SelectedIcType.Name);

        //                await shapeService.CopyShapeToProposedAsync(shapeToSave.Geometry, newShapeId, SelectedIcType.Name);
        //            }
        //        }

        //        // 6. Update database records with final status
        //        foreach (var subId in submissionIdMap.Values)
        //        {
        //            await submissionService.UpdateSubmissionCountsAsync(subId, goodCounts.GetValueOrDefault(subId, 0), dupCounts.GetValueOrDefault(subId, 0));
        //        }
        //        string finalStatus = (goodCount == 0 && dupCount > 0) ? "Duplicate" : "Migrated";
        //        await deliverableService.UpdateDeliverableStatusAsync(newDelId, finalStatus, "Pass");

        //        // 7. Save test results and send notification
        //        await testResultService.SaveTestResultsAsync(finalTestResult, newDelId);
        //        bool emailWasSent =  await notificationService.SendConfirmationEmailAsync(newDelId, finalTestResult, SelectedIcType.Name, outlookApp);

        //        if (!emailWasSent)
        //        {
        //            StatusMessage = "Operation canceled by user.";
        //            IsEmailActionEnabled = true; // Re-enable the UI
        //            return; // ABORT the finalization
        //        }


        //        // 8. Move the processed email
        //        outlookService.MoveEmailToFolder(outlookApp, _currentEmail.Emailid, _currentIcSetting.OutlookInboxFolderPath, _currentIcSetting.OutlookProcessedFolderPath);

        //        StatusMessage = $"Successfully finalized submission as {newDelId}.";
        //    }
        //    catch (Exception ex)
        //    {
        //        Log.RecordError("A critical error occurred during the finalization process.", ex, "FinalizeSubmissionAsync");
        //        ArcGIS.Desktop.Framework.Dialogs.MessageBox.Show("An error occurred during finalization. Please check the logs.", "Error");
        //        IsEmailActionEnabled = true; // Re-enable buttons on failure
        //        return; // Stop the process
        //    }
        //    finally
        //    {
        //        if (outlookApp != null)
        //        {
        //            System.Runtime.InteropServices.Marshal.ReleaseComObject(outlookApp);
        //        }
        //    }

        //    // 9. Advance to the next email
        //    if (_emailQueues.TryGetValue(SelectedIcType.Name, out var emailsToProcess) && emailsToProcess.Any())
        //    {
        //        emailsToProcess.RemoveAt(0);
        //    }
        //    await ProcessNextEmail();
        //}
        private async Task FinalizeSubmissionAsync(IcTestResult finalTestResult)
        {
            IsEmailActionEnabled = false;
            bool saveSucceeded = true;
            StatusMessage = "Finalizing submission...";
            Log.RecordMessage("Finalization process started.", BisLogMessageType.Note);

            var deliverableService = new Services.DeliverableService();
            var submissionService = new Services.SubmissionService();
            var shapeService = new Services.ShapeProcessingService(IcRules, Log);
            var notificationService = new Services.NotificationService();
            var outlookService = new Services.OutlookService();
            var testResultService = new Services.TestResultService();
            Outlook.Application outlookApp = null;

            string newDelId = null;
            try
            {
                var goodCounts = new Dictionary<string, int>();
                var dupCounts = new Dictionary<string, int>();
                outlookApp = new Outlook.Application();

                // 1. Create the main deliverable record
                newDelId = await deliverableService.CreateNewDeliverableRecordAsync(
                    "EMAIL", SelectedIcType.Name, CurrentPrefId, _currentEmail.ReceivedTime);
                CurrentDelId = newDelId;

                // 2. Update the RefId for all tests to use the new permanent ID
                finalTestResult.UpdateAllRefIds(newDelId);

                // 3. Save all metadata
                await deliverableService.UpdateEmailInfoRecordAsync(newDelId, _currentEmail, _currentClassification, _currentIcSetting.OutlookInboxFolderPath);
                await deliverableService.UpdateContactInfoRecordAsync(newDelId, _currentEmail);
                var bodyParser = new Services.EmailBodyParserService(SelectedIcType.Name);
                var bodyData = bodyParser.GetFieldsFromBody(_currentEmail.Body);
                await deliverableService.UpdateBodyDataRecordAsync(newDelId, bodyData);

                // 4. Record the submission filesets to get their IDs
                var submissionIdMap = await submissionService.RecordSubmissionsAsync(
                    newDelId, SelectedIcType.Name, _currentAttachmentAnalysis.IdentifiedFileSets);
                await submissionService.RecordPhysicalFilesAsync(newDelId, _currentAttachmentAnalysis.AllFiles, submissionIdMap);

                // 5. If the submission was approved, process and save the shapes
                if (finalTestResult.Passed)
                {
                    foreach (var shapeToSave in _selectedShapes)
                    {
                        StatusMessage = $"Processing shape {shapeToSave.ShapeReferenceId}...";
                        string submissionId = submissionIdMap.GetValueOrDefault(shapeToSave.SourceFile);
                        if (string.IsNullOrEmpty(submissionId))
                        {
                            Log.RecordError($"Could not find a submission ID for shape from file '{shapeToSave.SourceFile}'. Skipping shape record.", null, nameof(OnSave));
                            saveSucceeded = false;
                            continue;
                        }

                        string newShapeId = await shapeService.GetNextShapeIdAsync(newDelId, _currentIcSetting.IdPrefix);

                        bool recordCreated = await shapeService.RecordShapeInfoAsync(newShapeId, submissionId, newDelId, CurrentPrefId, SelectedIcType.Name);
                        if (!recordCreated)
                        {
                            Log.RecordError($"Aborting processing for this shape because its info record could not be created.", null, nameof(OnSave));
                            saveSucceeded = false;
                            continue;
                        }

                        bool isDuplicate = await shapeService.IsDuplicateInProposedAsync(shapeToSave.Geometry, CurrentPrefId, SelectedIcType.Name);

                        if (isDuplicate)
                        {
                            if (!dupCounts.ContainsKey(submissionId)) dupCounts[submissionId] = 0;
                            dupCounts[submissionId]++;
                            await shapeService.UpdateShapeInfoFieldAsync(newShapeId, "SHAPE_STATUS", "Duplicate", SelectedIcType.Name);
                        }
                        else
                        {
                            if (!goodCounts.ContainsKey(submissionId)) goodCounts[submissionId] = 0;
                            goodCounts[submissionId]++;
                            await shapeService.UpdateShapeInfoFieldAsync(newShapeId, "SHAPE_STATUS", "To Be Reviewed", SelectedIcType.Name);
                        }

                        await shapeService.UpdateShapeInfoFieldAsync(newShapeId, "CREATED_BY", "Crawler", SelectedIcType.Name);
                        await shapeService.UpdateShapeInfoFieldAsync(newShapeId, "CENTROID_X", shapeToSave.Geometry.Extent.Center.X, SelectedIcType.Name);
                        await shapeService.UpdateShapeInfoFieldAsync(newShapeId, "CENTROID_Y", shapeToSave.Geometry.Extent.Center.Y, SelectedIcType.Name);
                        await shapeService.UpdateShapeInfoFieldAsync(newShapeId, "SITE_DIST", shapeToSave.DistanceFromSite, SelectedIcType.Name);

                        bool copySuccess = await shapeService.CopyShapeToProposedAsync(shapeToSave.Geometry, newShapeId, SelectedIcType.Name);
                        if (!copySuccess)
                        {
                            saveSucceeded = false;
                            break;
                        }
                    }
                }


                // 6. Check our success flag. Only proceed if ALL shapes were saved correctly.
                if (!saveSucceeded)
                {
                    StatusMessage = "Save process halted due to an error copying a shape.";
                    // We do NOT proceed. The user can review the log and try saving again later.
                    // We should consider offering to roll back the database changes here in a future version.
                    return;
                }

                // 7. Move the original submission files to the final network archive location.
                //    This now happens for both passed and failed submissions to ensure a record is kept.
                if (_currentAttachmentAnalysis != null && _currentAttachmentAnalysis.IdentifiedFileSets != null)
                {
                    StatusMessage = "Archiving submission files...";
                    // A. Archive the files that came from the email
                    var emailFilesets = _currentAttachmentAnalysis.IdentifiedFileSets.Where(fs => fs.path != "manual").ToList();
                    if (emailFilesets.Any())
                    {
                        await submissionService.MoveAllSubmissionsAsync(
                            newDelId,
                            emailFilesets,
                            submissionIdMap,
                            _currentIcSetting.AsSubmittedPath);
                    }

                    // B. Export and archive the features from the "manually_added" layer
                    var manualFileset = _currentAttachmentAnalysis.IdentifiedFileSets.FirstOrDefault(fs => fs.path == "manual");
                    if (manualFileset != null && _manualAddLayer != null)
                    {
                        string submissionId = submissionIdMap.GetValueOrDefault(manualFileset.fileName);
                        if (!string.IsNullOrEmpty(submissionId))
                        {
                            string deliverableFolder = Path.Combine(_currentIcSetting.AsSubmittedPath, newDelId);
                            string submissionFolder = Path.Combine(deliverableFolder, submissionId);
                            Directory.CreateDirectory(submissionFolder);
                            string exportPath = Path.Combine(submissionFolder, "manually_added.shp");

                            var parameters = Geoprocessing.MakeValueArray(_manualAddLayer, exportPath);
                            var gpResult = await Geoprocessing.ExecuteToolAsync("management.CopyFeatures", parameters);

                            if (!gpResult.IsFailed)
                            {
                                Log.RecordMessage($"Successfully exported manual features to {exportPath}", BisLogMessageType.Note);
                                // Clear the layer after successful export and save
                                await QueuedTask.Run(() => {
                                    var editOp = new EditOperation() { Name = "Clear Manual Layer" };

                                    // 1. Get a cursor for all features in the layer's table.
                                    using (var cursor = _manualAddLayer.GetTable().Search(null, false))
                                    {
                                        // 2. Iterate through the cursor and collect all the Object IDs.
                                        var oids = new List<long>();
                                        while (cursor.MoveNext())
                                        {
                                            using (var row = cursor.Current)
                                            {
                                                oids.Add(row.GetObjectID());
                                            }
                                        }

                                        // 3. Queue the deletion of those specific OIDs.
                                        if (oids.Any())
                                        {
                                            editOp.Delete(_manualAddLayer, oids);
                                        }
                                    }

                                    // 4. Execute the operation.
                                    editOp.Execute();
                                });
                            }
                            else
                            {
                                Log.RecordError($"Failed to export manual features. Error: {string.Join(";", gpResult.Messages.Select(m => m.Text))}", null, "FinalizeSubmissionAsync");
                            }
                        }
                    }
                }

                // 8. Update database records with final status
                foreach (var subId in submissionIdMap.Values)
                {
                    await submissionService.UpdateSubmissionCountsAsync(subId, goodCounts.GetValueOrDefault(subId, 0), dupCounts.GetValueOrDefault(subId, 0));
                }
                string finalStatus = (goodCounts.Values.Sum() > 0) ? "Migrated" : "Failed";
                string finalValidity = finalTestResult.Passed ? "Pass" : "Fail";
                await deliverableService.UpdateDeliverableStatusAsync(newDelId, finalStatus, finalValidity);

                // 9. Save test results and send notification
                await testResultService.SaveTestResultsAsync(finalTestResult, newDelId);

                // **MODIFIED**: Pass the list of all submitted files to the email service
                var emailWasSent = await notificationService.SendConfirmationEmailAsync(
                    newDelId,
                    finalTestResult,
                    SelectedIcType.Name,
                    outlookApp,
                    _currentAttachmentAnalysis.AllFiles);

                if (!emailWasSent)
                {
                    StatusMessage = "Operation canceled by user.";
                    IsEmailActionEnabled = true; // Re-enable the UI
                                                 // Important: We need to reverse the database changes here or provide a manual way to clean up.
                                                 // For now, we will stop the process.
                    Log.RecordMessage($"User canceled email send for deliverable {newDelId}. The database record was created but the email was not moved. Manual cleanup may be required.",BisLogMessageType.FatalError);
                    return; // ABORT the finalization
                }

                // 10. Move the processed email
                var (store, folder) = OutlookService.ParseOutlookPath(_currentIcSetting.OutlookInboxFolderPath);
                outlookService.MoveEmailToFolder(outlookApp, _currentEmail.Emailid, $"\\\\{store}\\{folder}", _currentIcSetting.OutlookProcessedFolderPath);

                StatusMessage = $"Successfully finalized submission as {newDelId}.";
            }
            catch (Exception ex)
            {
                Log.RecordError("A critical error occurred during the finalization process.", ex, "FinalizeSubmissionAsync");
                ArcGIS.Desktop.Framework.Dialogs.MessageBox.Show("An error occurred during finalization. Please check the logs.", "Error");
                IsEmailActionEnabled = true; // Re-enable buttons on failure
                return; // Stop the process
            }
            finally
            {
                if (outlookApp != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(outlookApp);
                }
            }

            // 11. Advance to the next email
            if (_emailQueues.TryGetValue(SelectedIcType.Name, out var emailsToProcess) && emailsToProcess.Any())
            {
                emailsToProcess.RemoveAt(0);
            }
            await ProcessNextEmail();
        }

        #endregion


    }
}
