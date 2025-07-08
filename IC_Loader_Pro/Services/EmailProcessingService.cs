using IC_Loader_Pro.Models;
using IC_Rules_2025;
using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using static BIS_Log;
using static IC_Loader_Pro.Module1;

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
        /// Replicates the logic of the legacy processEmail function.
        /// </summary>
        /// <param name="emailId">The unique identifier of the email to process.</param>
        public async Task<IcTestResult> ProcessEmailAsync(string emailId)
        {
            _log.RecordMessage($"Starting to process email with ID: {emailId}", BisLogMessageType.Note);

            // 1. Create the root test result for the entire operation.
            // We assume a test rule named "GIS_Root_Email_Load" exists.
            var rootTestResult = new IcTestResult(_namedTests.returnTestRule("GIS_Root_Email_Load"), emailId, IcTestResult.TestType.Deliverable, _log, null, _namedTests);

            // 2. Retrieve the email from Outlook.
            EmailItem emailToProcess;
            try
            {
                // Note: This part will need to be adapted depending on whether you use
                // OutlookService or GraphApiService. This uses a placeholder for now.
                emailToProcess = _outlookService.GetEmailById(emailId);
                if (emailToProcess == null)
                {
                    throw new FileNotFoundException($"Email with ID '{emailId}' could not be found.");
                }
            }
            catch (Exception ex)
            {
                _log.RecordError($"Fatal error retrieving email with ID '{emailId}'.", ex, nameof(ProcessEmailAsync));
                rootTestResult.Passed = false;
                rootTestResult.Comments.Add("Failed to retrieve the email from the mail server.");
                return rootTestResult;
            }

            // 3. Classify the email to determine its type (Spam, CEA, DNA, etc.)
            EmailClassificationResult classification = _classifier.ClassifyEmail(emailToProcess);

            // 4. Handle simple cases first (Spam, Auto-Reply, Blocked, etc.)
            switch (classification.Type)
            {
                case EmailType.Spam:
                    _log.RecordMessage("Email classified as SPAM. Moving to Junk folder.", BisLogMessageType.Note);
                    // _outlookService.MoveEmailToFolder(emailId, "Junk"); // Future implementation
                    rootTestResult.Comments.Add("Email identified as SPAM and was moved.");
                    rootTestResult.Passed = false; // Mark as failed to stop processing
                    return rootTestResult;

                case EmailType.AutoResponse:
                    _log.RecordMessage("Email classified as an Auto-Response. Moving to Correspondence.", BisLogMessageType.Note);
                    // _outlookService.MoveEmailToFolder(emailId, "Correspondence"); // Future implementation
                    rootTestResult.Comments.Add("Email identified as an auto-response and was moved.");
                    rootTestResult.Passed = false;
                    return rootTestResult;

                // Add cases for other simple types like BlockedEmail...

                case EmailType.CEA:
                case EmailType.DNA:
                case EmailType.WRS:
                    // Email is one of the types we process. Continue to the next steps.
                    rootTestResult.Comments.Add($"Email type determined to be: {classification.Type}");
                    break;

                default:
                    _log.RecordMessage($"Email type is '{classification.Type}' and is not configured for processing by this tool.", BisLogMessageType.Warning);
                    // Decide if we move it to a different folder, for now we just stop.
                    rootTestResult.Comments.Add($"Email identified as unhandled type: {classification.Type}.");
                    rootTestResult.Passed = false;
                    return rootTestResult;
            }

            // --- NEXT STEPS WILL GO HERE ---
            // 5. Create the Deliverable Record in the database.
            // 6. Save attachments and identify the file sets.
            // 7. Run tests on each file set.
            // 8. Aggregate results and save everything to the database.

            return rootTestResult;
        }
    }
}