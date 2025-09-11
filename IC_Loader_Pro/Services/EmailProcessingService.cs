using ArcGIS.Core.Data.UtilityNetwork.Trace;
using ArcGIS.Core.Geometry;
using ArcGIS.Core.Internal.Threading.Tasks;
using BIS_Tools_DataModels_2025;
using IC_Loader_Pro.Models;
using IC_Rules_2025;
using Microsoft.Graph.Models;
using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime;
using System.Threading.Tasks;
using System.Windows.Documents;
using System.Windows.Media;
using static ArcGIS.Desktop.Internal.Mapping.Symbology.GlyphPickerViewModel;
using static BIS_Log;
using static IC_Loader_Pro.Module1;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace IC_Loader_Pro.Services
{
    /// <summary>
    /// Orchestrates the end-to-end process of receiving an email,
    /// classifying it, running tests, and recording the results.
    /// </summary>
    public class EmailProcessingService
    {
        private readonly BIS_Log _log;
        private readonly IC_Rules _rules;
        private readonly EmailClassifierService _classifier;
        private readonly OutlookService _outlookService; // or GraphApiService
        private readonly IcNamedTests _namedTests;

        public EmailProcessingService(IC_Rules rules, IcNamedTests namedTests, BIS_Log log)
        {
            _log = log;
            _rules = rules;
            _classifier = new EmailClassifierService(rules,_log); // It depends on the rules engine
            _outlookService = new OutlookService(); // Assuming we use this for now
            _namedTests = namedTests ?? throw new ArgumentNullException(nameof(namedTests));
        }

        /// <summary>
        /// The main entry point for processing a single email.
        /// This version receives pre-fetched and pre-classified email objects.
        /// </summary>
        /// <param name="emailToProcess">The complete EmailItem object, including body and attachments.</param>
        /// <param name="classification">The result of the initial classification.</param>
        /// /// <param name="sourceFolderPath">The Outlook folder path where the email currently resides.</param>
        /// <param name="sourceStoreName">The name of the Outlook store (mailbox) where the email resides.</param>
        /// <returns>The master test result for the entire operation.</returns>
        // The method's return type is changed to our new wrapper class
        // public async Task<EmailProcessingResult> ProcessEmailAsync(Outlook.Application outlookApp, EmailItem emailToProcess, EmailClassificationResult classification, string selectedIcType,
        //string sourceFolderPath,
        //string sourceStoreName,
        //bool wasManuallyClassified,
        //EmailType finalType,
        //Func<string, Task<MapPoint>> getSiteCoordsTask)
        // {
        //     _log.RecordMessage($"Starting to process email with ID: {emailToProcess.Emailid}", BisLogMessageType.Note);

        //     var rootTestResult = new IcTestResult(_namedTests.returnTestRule("GIS_Root_Email_Load"), "-1", IcTestResult.TestType.Deliverable, _log, null, _namedTests);
        //     var filesetTestResults = new List<IcTestResult>();
        //     var currentIcSetting = _rules.ReturnIcGisTypeSettings(selectedIcType);
        //     AttachmentAnalysisResult attachmentAnalysis = null;

        //     if (currentIcSetting == null)
        //     {
        //         rootTestResult.Passed = false;
        //         rootTestResult.Comments.Add($"Fatal error: Rules for queue '{selectedIcType}' not found.");
        //         return new EmailProcessingResult { TestResult = rootTestResult };
        //     }

        //     var subjectLineTest = _namedTests.returnNewTestResult("GIS_Subjectline_Tests", "-1", IcTestResult.TestType.Deliverable);
        //     if (wasManuallyClassified)
        //     {
        //         subjectLineTest.AddComment($"User manually classified email as '{finalType}'.");
        //     }
        //     subjectLineTest.Passed = !string.IsNullOrWhiteSpace(emailToProcess.Subject);
        //     subjectLineTest.AddComment(subjectLineTest.Passed ? "Subject line is present." : "Original subject was empty.");
        //     rootTestResult.AddSubordinateTestResult(subjectLineTest);

        //     var (prefIdTest, siteLocation) = await ValidatePrefIdAsync(classification, getSiteCoordsTask);
        //     rootTestResult.AddSubordinateTestResult(prefIdTest);

        //     var outlookService = new OutlookService();

        //     // 1. Handle Spam, Auto-Replies, and Mismatched IC Types
        //     if (finalType == EmailType.Spam || finalType == EmailType.AutoResponse || (finalType.Name != selectedIcType && !wasManuallyClassified))
        //     {
        //         string moveReason;
        //         string fullDestPath;

        //         if (finalType.Name != selectedIcType && !wasManuallyClassified)
        //         {
        //             var correctIcSetting = _rules.ReturnIcGisTypeSettings(finalType.Name);
        //             moveReason = $"Email type '{finalType.Name}' does not match the selected queue '{selectedIcType}'.";
        //             fullDestPath = correctIcSetting?.OutlookInboxFolderPath;
        //         }
        //         else
        //         {
        //             moveReason = $"Email identified as {finalType.Name}.";
        //             fullDestPath = finalType == EmailType.Spam ?
        //                 currentIcSetting.OutlookSpamFolderPath :
        //                 currentIcSetting.OutlookCorrespondenceFolderPath;
        //         }

        //         if (!string.IsNullOrEmpty(fullDestPath))
        //         {
        //             NotifyAndMoveEmail(outlookApp, emailToProcess, sourceFolderPath, sourceStoreName, fullDestPath, moveReason);
        //         }

        //         // Return a result with a null TestResult to signal an automatic advance.
        //         return new EmailProcessingResult { TestResult = null };
        //     }

        //     // 3. Process Attachments
        //     var attachmentService = new AttachmentService(this._rules, this._namedTests, Module1.FileTool, this._log);
        //     attachmentAnalysis = attachmentService.AnalyzeAttachments(emailToProcess.TempFolderPath, selectedIcType);
        //     if (attachmentAnalysis.TestResult != null)
        //     {
        //         filesetTestResults.Add(attachmentAnalysis.TestResult);
        //     }       


        //     if (attachmentAnalysis.TestResult.Comments.Contains("Email contains no attachments."))
        //     {
        //         rootTestResult.Passed = false; // The overall process is not a "pass"
        //         rootTestResult.Comments.Add("No attachments found. Treating as Correspondence.");

        //         string reason = "Email contains no attachments to process.";
        //         NotifyAndMoveEmail(outlookApp, emailToProcess.Emailid, sourceFolderPath, sourceStoreName, currentIcSetting.OutlookCorrespondenceFolderPath, reason, emailToProcess.Subject);

        //         return new EmailProcessingResult { TestResult = rootTestResult, AttachmentAnalysis = attachmentAnalysis };
        //     }

        //     if (attachmentAnalysis.TestResult.CumulativeAction.ResultAction != TestActionResponse.Pass)
        //     {
        //         HandleRejection(outlookApp, rootTestResult, currentIcSetting, sourceFolderPath, sourceStoreName);
        //         return new EmailProcessingResult { TestResult = rootTestResult, AttachmentAnalysis = attachmentAnalysis };
        //     }

        //     //  Handle No GIS Files Found
        //     if (!attachmentAnalysis.IdentifiedFileSets.Any())
        //     {
        //         rootTestResult.Passed = false;
        //         rootTestResult.Comments.Add("No valid GIS datasets found in attachments. Treating as Correspondence.");
        //         string reason = "No valid GIS datasets found in attachments.";
        //         NotifyAndMoveEmail(outlookApp, emailToProcess.Emailid, sourceFolderPath, sourceStoreName, currentIcSetting.OutlookCorrespondenceFolderPath, reason, emailToProcess.Subject);
        //         // outlookService.MoveEmailToFolder(outlookApp, emailToProcess.Emailid, sourceFolderPath, sourceStoreName, currentIcSetting.OutlookCorrespondenceFolderPath);
        //         return new EmailProcessingResult { TestResult = rootTestResult, AttachmentAnalysis = attachmentAnalysis };
        //     }

        //     var featureService = new FeatureProcessingService(_rules, _namedTests, _log);
        //     var featureProcessingContainerTest = _namedTests.returnNewTestResult("GIS_SubmissionFileCheck", emailToProcess.Emailid, IcTestResult.TestType.Submission);
        //     List<ShapeItem> foundShapes = await featureService.AnalyzeFeaturesFromFilesetsAsync(attachmentAnalysis.IdentifiedFileSets, selectedIcType, siteLocation, featureProcessingContainerTest);
        //     filesetTestResults.Add(featureProcessingContainerTest);
        //     _log.RecordMessage($"Successfully extracted and analyzed {foundShapes.Count} shape features.", BisLogMessageType.Note);
        //     if (!foundShapes.Any())
        //     {
        //         rootTestResult.Passed = false;
        //         rootTestResult.AddComment("The submission contained GIS files, but no readable polygon features were found within them.");
        //     }

        //     await Task.CompletedTask;

        //     return new EmailProcessingResult
        //     {
        //         TestResult = rootTestResult,
        //         AttachmentAnalysis = attachmentAnalysis,
        //         ShapeItems = foundShapes,
        //         FilesetTestResults = filesetTestResults
        //     };
        // }


        public async Task<EmailProcessingResult> ProcessEmailAsync(
         Outlook.Application outlookApp,
         EmailItem emailToProcess,
         EmailClassificationResult classification,
         string selectedIcType,
         string sourceFolderPath,
         string sourceStoreName,
         bool wasManuallyClassified,
         EmailType finalType,
         Func<string, Task<MapPoint>> getSiteCoordsTask)
        {
            _log.RecordMessage($"Starting to process email with ID: {emailToProcess.Emailid}", BisLogMessageType.Note);

            var rootTestResult = new IcTestResult(_namedTests.returnTestRule("GIS_Root_Email_Load"), "-1", IcTestResult.TestType.Deliverable, _log, null, _namedTests);
            // --- START OF NEW DUPLICATE FILENAME CHECK ---
            if (!string.IsNullOrEmpty(emailToProcess.ProcessingError))
            {
                var duplicateTest = _namedTests.returnNewTestResult("GIS_DuplicateFilenamesInAttachments", emailToProcess.Emailid, IcTestResult.TestType.Deliverable);
                duplicateTest.SetResult(false, emailToProcess.ProcessingError);
                rootTestResult.AddSubordinateTestResult(duplicateTest);
            }
            var currentIcSetting = _rules.ReturnIcGisTypeSettings(selectedIcType);
            AttachmentAnalysisResult attachmentAnalysis = null;

            if (currentIcSetting == null)
            {
                rootTestResult.Passed = false;
                rootTestResult.Comments.Add($"Fatal error: Rules for queue '{selectedIcType}' not found.");
                return new EmailProcessingResult { TestResult = rootTestResult };
            }

            var subjectLineTest = _namedTests.returnNewTestResult("GIS_Subjectline_Tests", "-1", IcTestResult.TestType.Deliverable);
            if (wasManuallyClassified)
            {
                subjectLineTest.AddComment($"User manually classified email as '{finalType}'.");
            }
            subjectLineTest.Passed = !string.IsNullOrWhiteSpace(emailToProcess.Subject);
            subjectLineTest.AddComment(subjectLineTest.Passed ? "Subject line is present." : "Original subject was empty.");
            rootTestResult.AddSubordinateTestResult(subjectLineTest);

            var (prefIdTest, siteLocation) = await ValidatePrefIdAsync(classification, getSiteCoordsTask);
            rootTestResult.AddSubordinateTestResult(prefIdTest);

            var outlookService = new OutlookService();

            if (finalType == EmailType.Spam || finalType == EmailType.AutoResponse || (finalType.Name != selectedIcType && !wasManuallyClassified))
            {
                string moveReason;
                string fullDestPath;

                if (finalType.Name != selectedIcType && !wasManuallyClassified)
                {
                    var correctIcSetting = _rules.ReturnIcGisTypeSettings(finalType.Name);
                    moveReason = $"Email type '{finalType.Name}' does not match the selected queue '{selectedIcType}'.";
                    fullDestPath = correctIcSetting?.OutlookInboxFolderPath;
                }
                else
                {
                    moveReason = $"Email identified as {finalType.Name}.";
                    fullDestPath = finalType == EmailType.Spam ?
                        currentIcSetting.OutlookSpamFolderPath :
                        currentIcSetting.OutlookCorrespondenceFolderPath;
                }

                if (!string.IsNullOrEmpty(fullDestPath))
                {
                    NotifyAndMoveEmail(outlookApp, emailToProcess, sourceFolderPath, sourceStoreName, fullDestPath, moveReason);
                }

                return new EmailProcessingResult { TestResult = null };
            }

            // Process Attachments and add their results directly to the main tree
            var attachmentService = new AttachmentService(this._rules, this._namedTests, Module1.FileTool, this._log);
            attachmentAnalysis = attachmentService.AnalyzeAttachments(emailToProcess.TempFolderPath, selectedIcType);
            rootTestResult.AddSubordinateTestResult(attachmentAnalysis.TestResult);

            if (attachmentAnalysis.TestResult.Comments.Contains("Email contains no attachments."))
            {
                rootTestResult.Passed = false;
                rootTestResult.Comments.Add("No attachments found. Treating as Correspondence.");
                string reason = "Email contains no attachments to process.";
                NotifyAndMoveEmail(outlookApp, emailToProcess, sourceFolderPath, sourceStoreName, currentIcSetting.OutlookCorrespondenceFolderPath, reason);
                return new EmailProcessingResult { TestResult = null, AttachmentAnalysis = attachmentAnalysis };
            }

            if (attachmentAnalysis.TestResult.CumulativeAction.ResultAction != TestActionResponse.Pass)
            {
                return new EmailProcessingResult { TestResult = rootTestResult, AttachmentAnalysis = attachmentAnalysis };
            }

            if (!attachmentAnalysis.IdentifiedFileSets.Any())
            {
                rootTestResult.Passed = false;
                rootTestResult.Comments.Add("No valid GIS datasets found in attachments. Treating as Correspondence.");
                string reason = "No valid GIS datasets found in attachments.";
                NotifyAndMoveEmail(outlookApp, emailToProcess, sourceFolderPath, sourceStoreName, currentIcSetting.OutlookCorrespondenceFolderPath, reason);
                return new EmailProcessingResult { TestResult = null, AttachmentAnalysis = attachmentAnalysis };
            }

            // Process Features and add their results directly to the main tree
            var featureService = new FeatureProcessingService(_rules, _namedTests, _log);
            var featureProcessingContainerTest = _namedTests.returnNewTestResult("GIS_SubmissionFileCheck", emailToProcess.Emailid, IcTestResult.TestType.Submission);
            List<ShapeItem> foundShapes = await featureService.AnalyzeFeaturesFromFilesetsAsync(attachmentAnalysis.IdentifiedFileSets, selectedIcType, siteLocation, featureProcessingContainerTest);
            rootTestResult.AddSubordinateTestResult(featureProcessingContainerTest);

            _log.RecordMessage($"Successfully extracted and analyzed {foundShapes.Count} shape features.", BisLogMessageType.Note);
            if (!foundShapes.Any())
            {
                rootTestResult.Passed = false;
                rootTestResult.AddComment("The submission contained GIS files, but no readable polygon features were found within them.");
            }

            await Task.CompletedTask;

            return new EmailProcessingResult
            {
                TestResult = rootTestResult,
                AttachmentAnalysis = attachmentAnalysis,
                ShapeItems = foundShapes
            };
        }










        /// <summary>
        /// Displays a notification to the user and then moves an email to a specified folder.
        /// </summary>
        private void NotifyAndMoveEmail(Outlook.Application outlookApp, EmailItem email, string sourceFolderPath, string sourceStoreName, string fullDestPath, string reason)
        {
            var (destStore, destFolder) = OutlookService.ParseOutlookPath(fullDestPath);

            // Show the notification to the user
            string message = $"The following email will be automatically moved:\n\n" +
                             $"Subject: {email.Subject}\n" +
                             $"Reason: {reason}\n" +
                             $"Destination: {destFolder}";

            ArcGIS.Desktop.Framework.Dialogs.MessageBox.Show(message, "Automated Email Move", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Information);

            string fullSourcePath = $"\\\\{sourceStoreName}\\{sourceFolderPath}";

            // Move the email
            _outlookService.MoveEmailToFolder(
                outlookApp,
                email.Emailid,
                fullSourcePath,
                fullDestPath
            );
        }

        // In Services/EmailProcessingService.cs

        /// <summary>
        /// (SHELL METHOD) Handles the final processing steps for a valid submission,
        /// including creating the deliverable record and saving all test results.
        /// </summary>
        /// <param name="rootTestResult">The root test result containing all sub-tests.</param>
        /// <param name="attachmentAnalysis">The results of the attachment analysis.</param>
        /// <returns>The newly generated Deliverable ID from the database.</returns>
        public async Task<string> FinalizeAndSaveAsync(IcTestResult rootTestResult, AttachmentAnalysisResult attachmentAnalysis)
        {
            _log.RecordMessage("Finalizing and saving submission...", BisLogMessageType.Note);

            // --- FUTURE LOGIC ---
            // 1. Create the Deliverable record in the database using the email info.
            //    This database call would return the new Deliverable ID.
            string newDeliverableId = "DEL-54321"; // Placeholder
            _log.RecordMessage($"Generated new Deliverable ID: {newDeliverableId}", BisLogMessageType.Note);

            // 2. Run final, detailed tests on the GIS filesets found in attachmentAnalysis.
            //    These would be added as more subordinate tests to rootTestResult.

            // 3. Save the entire test result hierarchy to the database.
            //rootTestResult.RecordResults(newDeliverableId); // Pass the new ID to link the tests
            _log.RecordMessage("Successfully (will be) saved all test results to the database.", BisLogMessageType.Note);

            // Make the method async
            await Task.CompletedTask;

            return newDeliverableId;
        }



        public void HandleRejection(Outlook.Application outlookApp, IcTestResult testResult,IcGisTypeSetting icSetting, string sourceFolderPath, string sourceStoreName)
        {
            _log.RecordMessage("Handling rejection...", BisLogMessageType.Note);

            // --- NEW: Step 1: Create the Deliverable Record and get the new ID ---
            // This is the same first step that a "Save" operation would perform.
            // In the future, this will be a real database call.
            string newDeliverableId = "DEL-" + new Random().Next(10000, 99999); // Placeholder for DB call
            _log.RecordMessage($"Generated new Deliverable ID for rejected submission: {newDeliverableId}", BisLogMessageType.Note);
            // --------------------------------------------------------------------

            // 2. Record the final test result hierarchy to the database.
            // We now pass the new Deliverable ID to link the tests to the record.
            //testResult.RecordResults(newDeliverableId);
            _log.RecordMessage("Rejection result (Will have) been recorded to the database.", BisLogMessageType.Note);

            // 3. (SHELL) Generate the content for the rejection email.
            var rejectionEmailBody = string.Join("\n", testResult.Comments);
            _log.RecordMessage($"Generated rejection email body:\n{rejectionEmailBody}", BisLogMessageType.Note);
            // In the future, you would use this to create and send an email.

            // 4. Move the email to the 'Proccessed' folder.
            // The RefId on the test result is the email's MessageId.
            string emailMessageId = testResult.RefId;
            var outlookService = new OutlookService();
            string fullSourcePath = $"\\\\{sourceStoreName}\\{sourceFolderPath}";
            string fullDestinationPath = icSetting.OutlookProcessedFolderPath;

            // 2. Call the correct 4-argument version of the method.
            outlookService.MoveEmailToFolder(
                outlookApp,
                emailMessageId,
                fullSourcePath,
                fullDestinationPath
            );
        }

        /// <summary>
        /// Validates the Preference IDs found in the email classification, checking for both
        /// validity and the existence of coordinates.
        /// </summary>
        /// <param name="classification">The result of the email classification.</param>
        /// <param name="getSiteCoordsTask">A function that can be called to get the site coordinates.</param>
        /// <returns>A tuple containing the detailed IcTestResult and the found MapPoint.</returns>
        private async Task<(IcTestResult TestResult, MapPoint SiteLocation)> ValidatePrefIdAsync(
            EmailClassificationResult classification,
            Func<string, Task<MapPoint>> getSiteCoordsTask)
        {
            IcTestResult prefIdTestResult;
            MapPoint foundLocation = null;

            if (!classification.PrefIds.Any())
            {
                // Case 1: No PrefID was found in the subject line.
                prefIdTestResult = _namedTests.returnNewTestResult("GIS_NoPrefIdInSubjectLine", "", IcTestResult.TestType.Deliverable);
                prefIdTestResult.Passed = false;
            }
            else if (classification.PrefIds.Count > 1)
            {
                // Case 2: Multiple potential PrefIDs were found. We must find exactly one valid ID.
                _log.RecordMessage($"Found multiple potential PrefIDs: {string.Join(", ", classification.PrefIds)}. Validating each...", BisLogMessageType.Note);

                var validPrefIds = new List<string>();
                foreach (var id in classification.PrefIds)
                {
                    if (await IsValidPrefIdInDbAsync(id))
                    {
                        validPrefIds.Add(id);
                    }
                }

                if (validPrefIds.Count == 1)
                {
                    // Success! We found exactly one valid ID.
                    string validId = validPrefIds.First();
                    _log.RecordMessage($"Found a single valid PrefID ('{validId}') among the candidates.", BisLogMessageType.Note);

                    // Now, we proceed with this single valid ID to check for coordinates.
                    foundLocation = await getSiteCoordsTask(validId);
                    if (foundLocation == null)
                    {
                        prefIdTestResult = _namedTests.returnNewTestResult("GIS_NoValidPrefIdInSubjectLine", "", IcTestResult.TestType.Deliverable);
                        prefIdTestResult.Comments.Add($"PrefID '{validId}' is valid but has no coordinates in the database.");
                        prefIdTestResult.Passed = false;
                    }
                    else
                    {
                        prefIdTestResult = _namedTests.returnNewTestResult("GIS_ValidPrefIdFound", "", IcTestResult.TestType.Deliverable);
                        prefIdTestResult.Comments.Add($"Found coordinates for single valid PrefID: {validId}");
                        prefIdTestResult.Passed = true;
                    }
                }
                else
                {
                    // Failure: We found zero or more than one valid ID. The submission is ambiguous.
                    prefIdTestResult = _namedTests.returnNewTestResult("GIS_MultipleValidPrefIDsOnSubjectLine", "", IcTestResult.TestType.Deliverable);
                    prefIdTestResult.Comments.Add(validPrefIds.Any() ?
                        $"Found multiple valid PrefIDs ({string.Join(", ", validPrefIds)}), making the submission ambiguous." :
                        "Found multiple potential IDs, but none were valid.");
                    prefIdTestResult.Passed = false;
                }
            }
            else
            {
                // Case 3: Exactly one PrefID was found.
                string prefId = classification.PrefIds.First();

                bool isIdValid = await IsValidPrefIdInDbAsync(prefId);

                if (!isIdValid)
                {
                    // The PrefID itself is not a valid identifier in the database.
                    prefIdTestResult = _namedTests.returnNewTestResult("GIS_NoValidPrefIdInSubjectLine", "", IcTestResult.TestType.Deliverable);
                    prefIdTestResult.Comments.Add($"The identifier '{prefId}' is not a valid PrefID.");
                    prefIdTestResult.Passed = false;
                }
                else
                {
                    // The ID is valid. NOW, check for coordinates.
                    foundLocation = await getSiteCoordsTask(prefId);
                    if (foundLocation == null)
                    {
                        // The PrefID is valid but has no matching coordinates.
                        prefIdTestResult = _namedTests.returnNewTestResult("GIS_NoValidPrefIdWithCoordsInSubjectLine", "", IcTestResult.TestType.Deliverable);
                        prefIdTestResult.Comments.Add($"PrefID '{prefId}' is valid but has no coordinates in the database.");
                        prefIdTestResult.Passed = false;
                    }
                    else
                    {
                        // Success! A single, valid PrefID with coordinates was found.
                        prefIdTestResult = _namedTests.returnNewTestResult("GIS_ValidPrefIdFound", "", IcTestResult.TestType.Deliverable);
                        prefIdTestResult.Comments.Add($"Found coordinates for PrefID: {prefId}");
                        prefIdTestResult.Passed = true;
                    }
                }
            }

            return (prefIdTestResult, foundLocation);
        }

        // In IC_Loader_Pro/Services/EmailProcessingService.cs

        /// <summary>
        /// Displays a notification to the user and then moves an email to a specified folder.
        /// </summary>
        private void NotifyAndMoveEmail(Outlook.Application outlookApp, string emailId, string sourceFolderPath, string sourceStoreName, string fullDestPath, string reason, string subject)
        {
            var (destStore, destFolder) = OutlookService.ParseOutlookPath(fullDestPath);

            // Show the notification to the user
            string message = $"The following email will be automatically moved:\n\n" +
                             $"Subject: {subject}\n" +
                             $"Reason: {reason}\n" +
                             $"Destination: {destFolder}";

            ArcGIS.Desktop.Framework.Dialogs.MessageBox.Show(message, "Automated Email Move", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Information);

            string fullSourcePath = $"\\\\{sourceStoreName}\\{sourceFolderPath}";

            // Move the email
            _outlookService.MoveEmailToFolder(
            outlookApp,
            emailId,
            fullSourcePath,
            fullDestPath
            );

            //_outlookService.MoveEmailToFolder(
            //    outlookApp,
            //    emailId,
            //    sourceFolderPath,
            //    sourceStoreName,
            //    destFolder
            //);
        }


        /// <summary>
        /// A helper that now calls the centralized validation method in the IC_Rules engine.
        /// </summary>
        private Task<bool> IsValidPrefIdInDbAsync(string prefId)
        {
            // The logic is now delegated to the rules engine.
            return Task.Run(() => _rules.IsValidPrefId(prefId));
        }
    }
}