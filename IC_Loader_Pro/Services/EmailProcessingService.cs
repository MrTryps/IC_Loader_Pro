using IC_Loader_Pro.Models;
using IC_Rules_2025;
using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime;
using System.Threading.Tasks;
using System.Windows.Media;
using static BIS_Log;
using static IC_Loader_Pro.Module1;
using BIS_Tools_DataModels_2025;

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
            EmailItem emailToProcess,
            EmailClassificationResult classification,
            string selectedIcType,
            string sourceFolderPath,
            string sourceStoreName,
            EmailType? manuallySelectedType)
        {
            _log.RecordMessage($"Starting to process email with ID: {emailToProcess.Emailid}", BisLogMessageType.Note);

            var rootTestResult = new IcTestResult(_namedTests.returnTestRule("GIS_Root_Email_Load"), emailToProcess.Emailid, IcTestResult.TestType.Deliverable, _log, null, _namedTests);
            var currentIcSetting = _rules.ReturnIcGisTypeSettings(selectedIcType);
            AttachmentAnalysisResult attachmentAnalysis = null; // To hold the analysis results

            if (currentIcSetting == null)
            {
                rootTestResult.Passed = false;
                rootTestResult.Comments.Add($"Fatal error: Rules for queue '{selectedIcType}' not found.");
                return new EmailProcessingResult { TestResult = rootTestResult };
            }

            EmailType finalType = manuallySelectedType ?? classification.Type;

            // --- Corrected Subject Line Test ---
            var subjectLineTest = _namedTests.returnNewTestResult("GIS_EmptySubjectline", "-1", IcTestResult.TestType.Deliverable);
            if (classification.Type == EmailType.EmptySubjectline)
            {
                subjectLineTest.Passed = false; // An empty subject is a failed test
                subjectLineTest.AddComment($"Original subject was empty. User manually classified as '{finalType}'.");
            }
            else
            {
                subjectLineTest.Passed = true;
                subjectLineTest.AddComment("Subject line is present.");
            }
            rootTestResult.AddSubordinateTestResult(subjectLineTest);
            // ------------------------------------

            var outlookService = new OutlookService();

            // 1. Handle Spam and Auto-Replies
            if (finalType == EmailType.Spam || finalType == EmailType.AutoResponse)
            {
                // ... (This logic is correct and remains the same)
                return new EmailProcessingResult { TestResult = rootTestResult };
            }

            // 2. Handle Mismatched IC Types
            if (!finalType.ToString().Equals(selectedIcType, StringComparison.OrdinalIgnoreCase))
            {
                // ... (This logic is correct and remains the same)
                return new EmailProcessingResult { TestResult = rootTestResult };
            }

            // 3. If we get here, the email is the correct type. Proceed with full processing.
            rootTestResult.Comments.Add($"Email type confirmed as: {finalType}. Proceeding with attachment analysis.");

            var attachmentService = new AttachmentService(this._rules, this._namedTests, Module1.FileTool, this._log);
            attachmentAnalysis = attachmentService.AnalyzeAttachments(emailToProcess.TempFolderPath, selectedIcType);
            rootTestResult.AddSubordinateTestResult(attachmentAnalysis.TestResult);

            if (!attachmentAnalysis.TestResult.Passed)
            {
                _log.RecordMessage("Processing stopped due to attachment analysis failure. Handling as a rejection.", BisLogMessageType.Warning);
                rootTestResult.Passed = false;
                rootTestResult.Comments.Add("Attachment processing failed.");

                // Call the same shared rejection handler
                HandleRejection(rootTestResult, currentIcSetting, sourceFolderPath, sourceStoreName);

                return new EmailProcessingResult { TestResult = rootTestResult, AttachmentAnalysis = attachmentAnalysis };
            }

            // 4. Handle case where no GIS filesets were found
            if (attachmentAnalysis.IdentifiedFileSets.Count == 0)
            {
                _log.RecordMessage("No valid GIS datasets found in attachments.", BisLogMessageType.Warning);
                rootTestResult.Passed = false;
                rootTestResult.Comments.Add("No valid GIS datasets found in attachments.");
                string destinationPath = currentIcSetting.OutlookProcessedFolderPath;
                var (destStore, destFolder) = OutlookService.ParseOutlookPath(destinationPath);
                outlookService.MoveEmailToFolder(emailToProcess.Emailid, sourceFolderPath, sourceStoreName, destFolder);
                return new EmailProcessingResult { TestResult = rootTestResult, AttachmentAnalysis = attachmentAnalysis };
            }

            _log.RecordMessage($"Found {attachmentAnalysis.IdentifiedFileSets.Count} valid GIS datasets in attachments.", BisLogMessageType.Note);

            await Task.CompletedTask;

            // At the end, return the complete result object
            return new EmailProcessingResult
            {
                TestResult = rootTestResult,
                AttachmentAnalysis = attachmentAnalysis
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



        public void HandleRejection(IcTestResult testResult,IcGisTypeSetting icSetting, string sourceFolderPath, string sourceStoreName)
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
                emailMessageId,
                sourceFolderPath,
                sourceStoreName,
                destFolder
            );
        }
    }
}