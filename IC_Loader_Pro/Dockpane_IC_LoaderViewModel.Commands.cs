using ArcGIS.Core.CIM;
using ArcGIS.Core.Geometry;
using ArcGIS.Desktop.Framework;
using ArcGIS.Desktop.Framework.Contracts;
using ArcGIS.Desktop.Framework.Threading.Tasks;
using ArcGIS.Desktop.Internal.Mapping.Ribbon;
using ArcGIS.Desktop.Mapping;
using ArcGIS.Desktop.Mapping.Events;
using BIS_Tools_DataModels_2025;
using IC_Loader_Pro.Models;
using IC_Loader_Pro.Services;
using IC_Rules_2025;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;
using System.Windows.Input;
using static BIS_Log;
using static IC_Loader_Pro.Module1;
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

        private async Task OnSave()
        {
            if (_currentEmailTestResult == null || !_selectedShapes.Any())
            {
                ArcGIS.Desktop.Framework.Dialogs.MessageBox.Show(
                    "There must be at least one shape in the 'Selected Shapes to Use' list to save.", "Save Error");
                return;
            }

            IsEmailActionEnabled = false;
            StatusMessage = "Saving... Please wait.";
            Log.RecordMessage("Save process started.", BisLogMessageType.Note);

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
                // === Step 1 & 2: Create Deliverable and Save Metadata ===
                StatusMessage = "Creating deliverable record...";
                newDelId = await deliverableService.CreateNewDeliverableRecordAsync(
                    "EMAIL", SelectedIcType.Name, CurrentPrefId, _currentEmail.ReceivedTime);
                CurrentDelId = newDelId;
                _currentEmailTestResult.RefId = newDelId;

                await deliverableService.UpdateEmailInfoRecordAsync(newDelId, _currentEmail, _currentClassification, _currentIcSetting.OutlookInboxFolderPath);
                await deliverableService.UpdateContactInfoRecordAsync(newDelId, _currentEmail);
                var bodyParser = new Services.EmailBodyParserService(SelectedIcType.Name);
                var bodyData = bodyParser.GetFieldsFromBody(_currentEmail.Body);
                await deliverableService.UpdateBodyDataRecordAsync(newDelId, bodyData);

                // === Step 3: Record Submissions (filesets) ===
                StatusMessage = "Recording submissions...";
                var submissionIdMap = await submissionService.RecordSubmissionsAsync(
                    newDelId, SelectedIcType.Name, _currentAttachmentAnalysis.IdentifiedFileSets);
                await submissionService.RecordPhysicalFilesAsync(newDelId, _currentAttachmentAnalysis.AllFiles, submissionIdMap);

                // === Step 4: Copy Approved Shapes to the 'Proposed' Feature Class ===
                int shapeCounter = 0; // Counter for the sequential letter
                foreach (var shapeToSave in _selectedShapes)
                {
                    StatusMessage = $"Copying shape {shapeToSave.ShapeReferenceId}...";

                    char shapeSuffix = (char)('A' + shapeCounter);
                    string newShapeId = $"{newDelId}_{SelectedIcType.Name}_{shapeSuffix}";
                    shapeCounter++;

                    await shapeService.CopyShapeToProposedAsync(shapeToSave.Geometry, newShapeId, SelectedIcType.Name);
                }

                // === Step 5: Finalize, Record Results, and Clean Up ===
                StatusMessage = "Finalizing records...";

                // Update the final status for the deliverable.
                await deliverableService.UpdateDeliverableStatusAsync(newDelId, "Migrated", "Pass");

                // Record the final, compiled test results to the database.
                await testResultService.SaveTestResultsAsync(_currentEmailTestResult, newDelId);

                // Send confirmation email (shelled).
                await notificationService.SendConfirmationEmailAsync(newDelId);

                // Move the processed email.
                StatusMessage = "Moving processed email...";
                outlookApp = new Outlook.Application();
                // The refactored method now expects the full source and destination paths.
                outlookService.MoveEmailToFolder(
                    outlookApp,
                    _currentEmail.Emailid,
                    _currentIcSetting.OutlookInboxFolderPath,    // Full source path
                    _currentIcSetting.OutlookProcessedFolderPath // Full destination path
                );

                StatusMessage = $"Successfully saved submission as {newDelId}.";
                SelectedIcType.PassedCount++;
            }
            catch (Exception ex)
            {
                Log.RecordError("A critical error occurred during the save process.", ex, nameof(OnSave));
                ArcGIS.Desktop.Framework.Dialogs.MessageBox.Show("An error occurred while saving. Please check the logs.", "Save Error");
                IsEmailActionEnabled = true;
                return;
            }
            finally
            {
                if (outlookApp != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(outlookApp);
                    outlookApp = null;
                }
            }

            // Advance to the next email
            if (_emailQueues.TryGetValue(SelectedIcType.Name, out var emailsToProcess) && emailsToProcess.Any())
            {
                emailsToProcess.RemoveAt(0);
            }
            await ProcessNextEmail();
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
        private async Task OnReject()
        {
            Log.RecordMessage("Reject button was clicked.", BisLogMessageType.Note);
            Outlook.Application outlookApp = null;          
            if (SelectedIcType == null || string.IsNullOrEmpty(CurrentEmailId))
            {
                StatusMessage = "Nothing to reject.";
                return;
            }

            // Disable the action buttons while processing
            IsEmailActionEnabled = false;
            StatusMessage = "Processing rejection...";

            try
            {
                outlookApp = new Outlook.Application();
                // 1. Create the final test result for the manual rejection.
                var namedTests = new IcNamedTests(Log, PostGreTool);
                var rejectionTestResult = namedTests.returnNewTestResult(
                    "GIS_Root_Email_Load",
                    CurrentEmailId,
                    IcTestResult.TestType.Deliverable
                );
                rejectionTestResult.Passed = false;
                rejectionTestResult.AddComment("Submission was manually rejected by the user.");

                // 2. Update the UI stats immediately.
                SelectedIcType.FailedCount++;

                // 3. Call the shared rejection logic in the service.
                var processingService = new EmailProcessingService(IcRules, namedTests, Log);

                // Get the source folder info needed to move the email.
                string fullOutlookPath = _currentIcSetting.OutlookInboxFolderPath;
                var (storeName, folderPath) = OutlookService.ParseOutlookPath(fullOutlookPath);

                // This call is now much cleaner and delegates the work.
                // It needs to run on a background thread to avoid freezing the UI.
                await QueuedTask.Run(() =>
                    processingService.HandleRejection(outlookApp,rejectionTestResult, _currentIcSetting, folderPath, storeName)
                );
            }
            catch (Exception ex)
            {
                Log.RecordError("An error occurred during the rejection process.", ex, nameof(OnReject));
                ArcGIS.Desktop.Framework.Dialogs.MessageBox.Show(
                    "An error occurred during the rejection process. Please check the logs.",
                    "Rejection Error",
                    System.Windows.MessageBoxButton.OK,
                    System.Windows.MessageBoxImage.Error);
            }
            finally {                 // Ensure the Outlook application is released properly.
                if (outlookApp != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(outlookApp);
                    outlookApp = null;
                }
            }

            // 4. Advance to the next email in the queue.
            if (_emailQueues.TryGetValue(SelectedIcType.Name, out var emailsToProcess) && emailsToProcess.Any())
            {
                emailsToProcess.RemoveAt(0);
            }
            await ProcessNextEmail();
        }
        private async Task OnShowNotes()
        {
            Log.RecordMessage("Menu: Notes was clicked.", BisLogMessageType.Note);
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

        private void ShowTestResultWindow(IcTestResult testResult)
        {
            if (testResult == null)
            {
                ArcGIS.Desktop.Framework.Dialogs.MessageBox.Show("No test results are available to display.", "Show Results");
                return;
            }

            // This logic is moved from OnShowResults
            var testResultViewModel = new ViewModels.TestResultViewModel(testResult);
            var testResultWindow = new Views.TestResultWindow
            {
                DataContext = testResultViewModel,
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
                // Pass the site's geometry to our generic zoom helper.
                await ZoomToGeometryAsync(new List<Geometry> { _currentSiteLocation });
            }
        }

        #endregion


    }
}
