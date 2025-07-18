using ArcGIS.Desktop.Framework;
using ArcGIS.Desktop.Framework.Contracts;
using ArcGIS.Desktop.Framework.Threading.Tasks;
using BIS_Tools_DataModels_2025;
using IC_Loader_Pro.Services;
using IC_Rules_2025;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Input;
using static BIS_Log;
using static IC_Loader_Pro.Module1;


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

            // 1. Update stats and disable UI buttons.
            IsEmailActionEnabled = false;
            StatusMessage = "Finalizing and saving to database...";
            SelectedIcType.PassedCount++;

            try
            {
                // 2. Call the finalization service.
                var namedTests = new IcNamedTests(Log, PostGreTool);
                var processingService = new EmailProcessingService(IcRules, namedTests, Log);

                string newDelId = await processingService.FinalizeAndSaveAsync(_currentEmailTestResult, _currentAttachmentAnalysis);

                // Update the UI with the new Deliverable ID.
                CurrentDelId = newDelId;
                StatusMessage = "Successfully saved submission.";

                // 3. Move the processed email.
                var icSetting = IcRules.ReturnIcGisTypeSettings(SelectedIcType.Name);
                var outlookService = new OutlookService();
                string fullOutlookPath = icSetting.OutlookInboxFolderPath;
                var (storeName, folderPath) = OutlookService.ParseOutlookPath(fullOutlookPath);

                outlookService.MoveEmailToFolder(
                    CurrentEmailId,
                    folderPath,
                    storeName,
                    icSetting.OutlookProcessedFolderPath
                );
            }
            catch (Exception ex)
            {
                Log.RecordError("An error occurred during the save process.", ex, nameof(OnSave));
                ArcGIS.Desktop.Framework.Dialogs.MessageBox.Show(
                    "An error occurred while saving the submission. Please check the logs.",
                    "Save Error");
                // If save fails, roll back the "Passed" count.
                SelectedIcType.PassedCount--;
            }

            // 4. Advance to the next email.
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
            UpdateQueueStats(skipTestResult);

            // 3. Advance to the next email.
            await ProcessNextEmail();
        }
        private async Task OnReject()
        {
            Log.RecordMessage("Reject button was clicked.", BisLogMessageType.Note);

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
                var icSetting = IcRules.ReturnIcGisTypeSettings(SelectedIcType.Name);
                string fullOutlookPath = icSetting.OutlookInboxFolderPath;
                var (storeName, folderPath) = OutlookService.ParseOutlookPath(fullOutlookPath);

                // This call is now much cleaner and delegates the work.
                // It needs to run on a background thread to avoid freezing the UI.
                await QueuedTask.Run(() =>
                    processingService.HandleRejection(rejectionTestResult, icSetting, folderPath, storeName)
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

            // 4. Advance to the next email in the queue.
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

        #endregion
    }
}
