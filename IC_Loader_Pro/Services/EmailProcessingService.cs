using BIS_Tools_DataModels_2025;
using IC_Loader_Pro.Models;
using IC_Rules_2025;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime;
using System.Threading.Tasks;
using System.Windows.Media;
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
        public async Task<EmailProcessingResult> ProcessEmailAsync(
       Outlook.Application outlookApp,
       EmailItem emailToProcess,
       EmailClassificationResult classification,
       string selectedIcType,
       string sourceFolderPath,
       string sourceStoreName,
       bool wasManuallyClassified,
       EmailType finalType)
        {
            _log.RecordMessage($"Starting to process email with ID: {emailToProcess.Emailid}", BisLogMessageType.Note);

            var rootTestResult = new IcTestResult(_namedTests.returnTestRule("GIS_Root_Email_Load"), emailToProcess.Emailid, IcTestResult.TestType.Deliverable, _log, null, _namedTests);
            var currentIcSetting = _rules.ReturnIcGisTypeSettings(selectedIcType);
            AttachmentAnalysisResult attachmentAnalysis = null;

            if (currentIcSetting == null)
            {
                rootTestResult.Passed = false;
                rootTestResult.Comments.Add($"Fatal error: Rules for queue '{selectedIcType}' not found.");
                return new EmailProcessingResult { TestResult = rootTestResult };
            }

            var subjectLineTest = _namedTests.returnNewTestResult("GIS_Email_Submission_Tests", "-1", IcTestResult.TestType.Deliverable);
            if (wasManuallyClassified)
            {
                subjectLineTest.AddComment($"User manually classified email as '{finalType}'.");
            }
            subjectLineTest.Passed = !string.IsNullOrWhiteSpace(emailToProcess.Subject);
            subjectLineTest.AddComment(subjectLineTest.Passed ? "Subject line is present." : "Original subject was empty.");
            rootTestResult.AddSubordinateTestResult(subjectLineTest);

            var outlookService = new OutlookService();

            // 1. Handle Spam and Auto-Replies
            if (finalType == EmailType.Spam || finalType == EmailType.AutoResponse)
            {
                string fullDestPath = finalType == EmailType.Spam ?
                currentIcSetting.OutlookSpamFolderPath :
                currentIcSetting.OutlookCorrespondenceFolderPath;
                var (destStore, destFolder) = OutlookService.ParseOutlookPath(fullDestPath);
                bool moveSucceeded = outlookService.MoveEmailToFolder(
           outlookApp,
           emailToProcess.Emailid,
           sourceFolderPath,
           sourceStoreName,
           destFolder);
                rootTestResult.Passed = false;
                return new EmailProcessingResult { TestResult = rootTestResult };
            }

            // 2. Handle Mismatched IC Types
            if (finalType.Name != selectedIcType)
            {
                var correctIcSetting = _rules.ReturnIcGisTypeSettings(finalType.Name);
                if (correctIcSetting != null)
                {
                    outlookService.MoveEmailToFolder(outlookApp, emailToProcess.Emailid, sourceFolderPath, sourceStoreName, correctIcSetting.OutlookInboxFolderPath);
                }
                rootTestResult.Passed = false;
                return new EmailProcessingResult { TestResult = rootTestResult };
            }

            // 3. Process Attachments
            var attachmentService = new AttachmentService(this._rules, this._namedTests, Module1.FileTool, this._log);
            attachmentAnalysis = attachmentService.AnalyzeAttachments(emailToProcess.TempFolderPath, selectedIcType);
            rootTestResult.AddSubordinateTestResult(attachmentAnalysis.TestResult);

            if (!attachmentAnalysis.TestResult.Passed)
            {
                rootTestResult.Passed = false;
                HandleRejection(outlookApp, rootTestResult, currentIcSetting, sourceFolderPath, sourceStoreName);
                return new EmailProcessingResult { TestResult = rootTestResult, AttachmentAnalysis = attachmentAnalysis };
            }

            //  Handle No GIS Files Found
            if (!attachmentAnalysis.IdentifiedFileSets.Any())
            {
                rootTestResult.Passed = false;
                rootTestResult.Comments.Add("No valid GIS datasets found in attachments. Treating as Correspondence.");
                outlookService.MoveEmailToFolder(outlookApp, emailToProcess.Emailid, sourceFolderPath, sourceStoreName, currentIcSetting.OutlookCorrespondenceFolderPath);
                return new EmailProcessingResult { TestResult = rootTestResult, AttachmentAnalysis = attachmentAnalysis };
            }

            var featureService = new FeatureProcessingService(_rules, _namedTests, _log);
            List<ShapeItem> foundShapes = await featureService.AnalyzeFeaturesFromFilesetsAsync(attachmentAnalysis.IdentifiedFileSets, selectedIcType);
            _log.RecordMessage($"Successfully extracted and analyzed {foundShapes.Count} shape features.", BisLogMessageType.Note);            

            await Task.CompletedTask;

            return new EmailProcessingResult
            {
                TestResult = rootTestResult,
                AttachmentAnalysis = attachmentAnalysis,
                ShapeItems = foundShapes
            };
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
            string destinationPath = icSetting.OutlookProcessedFolderPath;
            var (destStore, destFolder) = OutlookService.ParseOutlookPath(destinationPath);
            outlookService.MoveEmailToFolder(
                outlookApp,
                emailMessageId,
                sourceFolderPath,
                sourceStoreName,
                destFolder
            );
        }
    }
}