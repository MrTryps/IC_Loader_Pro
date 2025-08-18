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

        #endregion

        #region Command Methods
        private async Task OnSave()
        {
            Log.RecordMessage("Save button was clicked.", BisLogMessageType.Note);

            if (_currentEmailTestResult == null)
            {
                ArcGIS.Desktop.Framework.Dialogs.MessageBox.Show("There is no processed email result to save.", "Save Error");
                return;
            }

            IsEmailActionEnabled = false; // Disable all action buttons
            StatusMessage = "Creating deliverable record in the database...";

            string newDelId = null;
            try
            {
                // 1. Create an instance of our new service.
                var deliverableService = new Services.DeliverableService();

                // 2. Call the method to create the record and get the new ID.
                //    We assume the source is an email for now.
                newDelId = await deliverableService.CreateNewDeliverableRecordAsync("EMAIL");

                // 3. Update the UI with the new ID.
                CurrentDelId = newDelId;
                StatusMessage = $"Successfully created Deliverable ID: {newDelId}";
                Log.RecordMessage(StatusMessage, BisLogMessageType.Note);

                if (Module1.IsInTestMode)
                {
                    var notesService = new Services.NotesService();
                    await notesService.RecordNoteAsync(newDelId, "This is a test deliverable created in Test Mode.", "Automation Note");
                    Log.RecordMessage($"Recorded 'Test Mode' note for deliverable {newDelId}.", BisLogMessageType.Note);
                }

                // --- FUTURE STEPS WILL GO HERE ---

                // TODO: Record contact info
                // TODO:
                // TODO: Save the IcTestResult hierarchy to the database, linked to newDelId.
                // TODO: Save the selected ShapeItems to the final feature class, linked to newDelId.
                // TODO: Move the processed email to the 'Processed' folder.
                // ---
            }
            catch (Exception ex)
            {
                Log.RecordError("An error occurred while creating the new deliverable record.", ex, nameof(OnSave));
                ArcGIS.Desktop.Framework.Dialogs.MessageBox.Show(
                    "An error occurred while saving the submission. Please check the logs.",
                    "Save Error");
                StatusMessage = "Save failed. Please review logs.";
                IsEmailActionEnabled = true; // Re-enable buttons on failure
                return; // Stop the save process
            }

            // If the save was successful, update the stats and advance to the next email.
            SelectedIcType.PassedCount++;

            if (_emailQueues.TryGetValue(SelectedIcType.Name, out var emailsToProcess) && emailsToProcess.Any())
            {
                emailsToProcess.RemoveAt(0);
            }
            await ProcessNextEmail();
        }

        // Also update OnSkip and OnReject
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
