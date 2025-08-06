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
                        emailsInQueue = await graphService.GetEmailsFromFolderPathAsync(outlookApp,outlookFolderPath, testSender, testModeFlag);
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

                if (emailToProcess == null)
                {
                    shouldAutoAdvance = true; // Mark for advancement
                    ArcGIS.Desktop.Framework.Dialogs.MessageBox.Show($"Could not retrieve the email: '{currentEmailSummary.Subject}'. It will be skipped.", "Email Retrieval Error");
                    return; // Exit the try block; the finally block will handle the rest.
                }

                var classification = new EmailClassifierService(IcRules, Log).ClassifyEmail(emailToProcess);

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

                _currentSiteLocation = await GetSiteCoordinatesAsync(CurrentPrefId);

                var processingService = new EmailProcessingService(IcRules, namedTests, Log);
                EmailProcessingResult processingResult = await processingService.ProcessEmailAsync(outlookApp, emailToProcess, classification, SelectedIcType.Name, folderPath, storeName, classification.WasManuallyClassified, finalEmailType, GetSiteCoordinatesAsync);

                _currentEmailTestResult = processingResult.TestResult;
               // UpdateQueueStats(_currentEmailTestResult);

                if (!_currentEmailTestResult.Passed)
                {
                    shouldAutoAdvance = true; // Mark for advancement
                    ShowTestResultWindow(_currentEmailTestResult);
                    return; // Exit try block
                }

                // --- On Success, Populate UI and Wait for User Input ---
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

                if (processingResult.ShapeItems?.Any() == true)
                {
                    await RunOnUIThread(() =>
                    {
                        _shapesToReview.Clear();
                        foreach (var shape in processingResult.ShapeItems)
                        {
                            _shapesToReview.Add(shape);
                        }
                    });
                    await RedrawAllShapesOnMapAsync();
                    await ZoomToAllAndSiteAsync();
                    }

                }
                StatusMessage = "Ready for review.";
                IsEmailActionEnabled = true;
            }
            catch (Exception ex)
            {
                shouldAutoAdvance = true; // Also advance on unexpected errors
                Log.RecordError($"An unexpected error occurred while processing email ID {currentEmailSummary.Emailid}", ex, "ProcessSelectedQueueAsync");
                ArcGIS.Desktop.Framework.Dialogs.MessageBox.Show("An unexpected error occurred. The application will advance to the next email.", "Processing Error");
            }
            finally
            {               
                CleanupTempFolder(emailToProcess);
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

            // 2. The rest of the method uses QueuedTask to draw the graphics.
            await QueuedTask.Run(() =>
            {
                var mapView = MapView.Active;
                if (mapView == null) return;

                var graphicsLayer = mapView.Map.FindLayers("IC Loader Shapes").FirstOrDefault() as GraphicsLayer;
                if (graphicsLayer == null) return;

                graphicsLayer.RemoveElements();

                foreach (var shapeItem in _shapesToReview)
                {
                    if (shapeItem.Geometry != null)
                    {
                         graphicsLayer.AddElement(shapeItem.Geometry, reviewSymbol, shapeItem.ShapeReferenceId.ToString());
                        //var newElement = ElementFactory.Instance.CreateGraphicElement(graphicsLayer, shapeItem.Geometry, reviewSymbol, shapeItem.ShapeReferenceId.ToString());
                        //var attributes = new Dictionary<string, object>
                        //{
                        //    { "ShapeRefID", shapeItem.ShapeReferenceId }
                        //};
                        //var graphic = new CIMPolygonGraphic
                        //{
                        //    Polygon = shapeItem.Geometry,
                        //    Symbol = reviewSymbol.MakeSymbolReference(), // Use MakeSymbolReference()
                        //    Attributes = attributes
                        //};
                        //graphicsLayer.AddElement(graphic);
                    }
                }

                foreach (var shapeItem in _selectedShapes)
                {
                    if (shapeItem.Geometry != null)
                    {
                         graphicsLayer.AddElement(shapeItem.Geometry, useSymbol, shapeItem.ShapeReferenceId.ToString());
                        //var newElement = ElementFactory.Instance.CreateGraphicElement(graphicsLayer, shapeItem.Geometry, useSymbol, shapeItem.ShapeReferenceId.ToString());
                        //var attributes = new Dictionary<string, object>
                        //{
                        //    { "ShapeRefID", shapeItem.ShapeReferenceId }
                        //};
                        //var graphic = new CIMPolygonGraphic
                        //{
                        //    Polygon = shapeItem.Geometry,
                        //    Symbol = useSymbol.MakeSymbolReference(), // Use MakeSymbolReference()
                        //    Attributes = attributes
                        //};
                        //graphicsLayer.AddElement(graphic);
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
        private async Task<MapPoint> GetSiteCoordinatesAsync(string prefId)
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



    }
}