// In IC_Loader_Pro/Services/NotificationService.cs

using ArcGIS.Desktop.Framework;
using BIS_Tools_DataModels_2025;
using IC_Loader_Pro.Models;
using IC_Rules_2025;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using static BIS_Log;
using static IC_Loader_Pro.Module1;

namespace IC_Loader_Pro.Services
{
    public class NotificationService
    {
        public async Task<bool> SendConfirmationEmailAsync(string deliverableId, IcTestResult testResult, string icType, Microsoft.Office.Interop.Outlook.Application outlookApp, List<AnalyzedFile> submittedFiles)
        {
            var emailToEdit = BuildReplyEmail(deliverableId, testResult, icType, submittedFiles);
            if (emailToEdit == null)
            {
                Log.RecordMessage("Email not built because no recipients were specified.", BisLogMessageType.Note);
                return true;
            }

            bool userClickedSend = false;
            await FrameworkApplication.Current.Dispatcher.InvokeAsync(() =>
            {
                var previewViewModel = new ViewModels.EmailPreviewViewModel(emailToEdit, testResult);
                var previewWindow = new Views.EmailPreviewWindow
                {
                    DataContext = previewViewModel,
                    Owner = FrameworkApplication.Current.MainWindow
                };

                if (previewWindow.ShowDialog() == true)
                {
                    userClickedSend = true;
                }
            });

            if (userClickedSend)
            {
                await SendOutlookEmailAsync(emailToEdit, outlookApp);
            }
            else
            {
                Log.RecordMessage("Email send was canceled by the user from the preview window.", BisLogMessageType.Note);
            }

            return userClickedSend;
        }
        private OutgoingEmail BuildReplyEmail(string deliverableId, IcTestResult testResult, string icType, List<AnalyzedFile> submittedFiles)
        {
            var finalCumulativeAction = testResult.CumulativeAction;

            if (finalCumulativeAction.EmailRp != true && finalCumulativeAction.EmailHazsite != true)
            {
                return null;
            }

            var outgoingEmail = new OutgoingEmail();
            var deliverableInfo = IcRules.ReturnEmailDeliverableInfo(deliverableId);
            var namedTests = new IcNamedTests(Log, PostGreTool);

            if (finalCumulativeAction.EmailRp == true && !string.IsNullOrEmpty(deliverableInfo.SenderEmail))
            {
                outgoingEmail.ToRecipients.Add(deliverableInfo.SenderEmail);
            }
            if (finalCumulativeAction.EmailHazsite == true)
            {
                outgoingEmail.BccRecipients.Add("SRPGIS@dep.nj.gov");
            }

            if (!outgoingEmail.ToRecipients.Any() && !outgoingEmail.BccRecipients.Any())
            {
                return null;
            }

            var parameters = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                { "DELID", deliverableId },
                { "PREFID", testResult.OutputParams.GetValueOrDefault("prefid", "N/A") },
                { "SUBJECTLINE", deliverableInfo.SubjectLine },
                { "SENDDATE", deliverableInfo.SubmitDate },
                { "Ic_Type", icType }
            };

            var rootRule = testResult.TestRule;
            string subjectTemplateText;
            string bodyTemplateText = "<P><div align='right'>{Template_GIS_OutgoingEmailID_Small}</DIV></P>";

            if (!testResult.Passed || finalCumulativeAction.ResultAction == TestActionResponse.Fail)
            {
                subjectTemplateText = rootRule.FailSubject?.ReplacementText ?? "Submission Processing Issue for {PREFID}";
                bodyTemplateText += rootRule.FailMessage?.ReplacementText ?? "<p>Your submission could not be processed due to the following reason(s):</p>";

                // --- START OF NEW, SIMPLIFIED LOGIC ---
                // 1. Get a simple, flat list of every specific test that failed.
                var allFailures = new List<IcTestResult>();
                FlattenFailedTests(testResult, allFailures);

                // 2. Build a unique message for each failure.
                var failureMessages = new List<string>();
                foreach (var failure in allFailures)
                {
                    string reasonTemplate = failure.TestRule.FailMessage?.ReplacementText ?? failure.TestRule.ErrorComment ?? "";
                    string runtimeComments = string.Join(" ", failure.Comments);
                    string finalReason = $"{reasonTemplate} {runtimeComments}".Trim();

                    if (!string.IsNullOrEmpty(finalReason))
                    {
                        var filledResult = namedTests.FillAllParameters(finalReason, parameters);
                        if (!string.IsNullOrWhiteSpace(filledResult.ProcessedText))
                        {
                            failureMessages.Add(filledResult.ProcessedText);
                        }
                    }
                }

                // 3. Format that list into clean HTML bullet points.
                if (failureMessages.Any())
                {
                    var formattedReasons = failureMessages.Distinct().Select(reason => $"<li>{reason}</li>");
                    outgoingEmail.AddToMainBody($"<ul>{string.Join("", formattedReasons)}</ul>");
                }
                // --- END OF NEW, SIMPLIFIED LOGIC ---

                outgoingEmail.AddToClosingText("{Template_GIS_Rejected}");
            }
            else
            {
                subjectTemplateText = rootRule.PassSubject?.ReplacementText ?? "Submission Processed for {PREFID}";
                bodyTemplateText += rootRule.PassMessage?.ReplacementText ?? "<p>Your submission has been processed successfully.</p>";
            }

            if (finalCumulativeAction.IncludeSubmittedFiles && submittedFiles != null)
            {
                foreach (var file in submittedFiles)
                {
                    string fullPath = Path.Combine(file.CurrentPath, file.FileName);
                    if (File.Exists(fullPath))
                    {
                        outgoingEmail.Attachments.Add(fullPath);
                    }
                }
            }

            var subjectResult = namedTests.FillAllParameters(subjectTemplateText, parameters);
            var openingTextResult = namedTests.FillAllParameters(bodyTemplateText, parameters);
            var mainBodyResult = namedTests.FillAllParameters(string.Join("", outgoingEmail.MainBodyText), parameters);
            var closingTextResult = namedTests.FillAllParameters(string.Join("", outgoingEmail.ClosingText), parameters);

            outgoingEmail.Subject = subjectResult.ProcessedText;
            outgoingEmail.OpeningText.Clear();
            outgoingEmail.AddToOpeningText(openingTextResult.ProcessedText);
            outgoingEmail.MainBodyText.Clear();
            outgoingEmail.AddToMainBody(mainBodyResult.ProcessedText);
            outgoingEmail.ClosingText.Clear();
            outgoingEmail.AddToClosingText(closingTextResult.ProcessedText);

            // ... (parameter checking logic remains the same)
            return outgoingEmail;
        }

        /// <summary>
        /// Recursively traverses a test result tree and creates a flat list of all failed tests
        /// that are not simply containers for other failed tests.
        /// </summary>
        private void FlattenFailedTests(IcTestResult testResult, List<IcTestResult> flatList)
        {
            // If the current test failed...
            if (!testResult.Passed)
            {
                // ...and it's a "leaf failure" (no children failed), add it to the list.
                // This captures the most specific error reason.
                if (!testResult.SubTestResults.Any(st => !st.Passed))
                {
                    flatList.Add(testResult);
                }
                else
                {
                    // If it's not a leaf failure, ignore this parent container and
                    // check its children for the specific errors.
                    foreach (var subResult in testResult.SubTestResults)
                    {
                        FlattenFailedTests(subResult, flatList);
                    }
                }
            }
        }


        /// <summary>
        /// Recursively traverses a test result tree and collects only the "leaf" failure messages.
        /// A leaf failure is a failed test that has no failed children, making it the most specific reason for the error.
        /// </summary>
        private void CollectLeafFailureMessages(IcTestResult testResult, List<string> messages, IcNamedTests namedTests, Dictionary<string, string> parameters)
        {
            // If the current test failed...
            if (!testResult.Passed)
            {
                // ...and it has NO children that also failed (meaning it's a root cause)...
                bool isLeafFailure = !testResult.SubTestResults.Any(st => !st.Passed);

                if (isLeafFailure)
                {
                    // ...then generate its message and add it to our list.
                    string failureReasonTemplate = testResult.TestRule.FailMessage?.ReplacementText ?? testResult.TestRule.ErrorComment ?? "";
                    string runtimeComments = string.Join(" ", testResult.Comments);
                    string finalReason = $"{failureReasonTemplate} {runtimeComments}".Trim();

                    if (!string.IsNullOrEmpty(finalReason))
                    {
                        var filledResult = namedTests.FillAllParameters(finalReason, parameters);
                        if (!string.IsNullOrWhiteSpace(filledResult.ProcessedText))
                        {
                            messages.Add(filledResult.ProcessedText);
                        }
                    }
                }
                else
                {
                    // This test failed because of a child. Ignore it and look deeper.
                    foreach (var subResult in testResult.SubTestResults)
                    {
                        CollectLeafFailureMessages(subResult, messages, namedTests, parameters);
                    }
                }
            }
        }

        /// <summary>
        /// Recursively traverses a test result hierarchy to build a formatted HTML list of failure messages
        /// using the 'FailMessage' or 'ErrorComment' from each failed test rule.
        /// </summary>
        private void BuildRejectionBody(IcTestResult testResult, StringBuilder failureMessages, IcNamedTests namedTests, Dictionary<string, string> parameters)
        {
            // If the test itself failed, add its reason to the list.
            if (!testResult.Passed)
            {
                string failureReasonTemplate = "";

                // Prioritize the specific 'FailMessage' template if it exists and is not just a placeholder.
                if (testResult.TestRule.FailMessage != null && !string.IsNullOrEmpty(testResult.TestRule.FailMessage.ReplacementText))
                {
                    failureReasonTemplate = testResult.TestRule.FailMessage.ReplacementText;
                }
                // Otherwise, fall back to the generic 'ErrorComment'.
                else if (!string.IsNullOrEmpty(testResult.TestRule.ErrorComment))
                {
                    failureReasonTemplate = testResult.TestRule.ErrorComment;
                }

                // Also include any specific comments added during runtime, as they often contain critical details.
                string runtimeComments = string.Join(" ", testResult.Comments);


                // If we found a reason, fill in its parameters and add it as a list item.
                if (!string.IsNullOrEmpty(failureReasonTemplate) || !string.IsNullOrEmpty(runtimeComments))
                {
                    string finalReason = $"{failureReasonTemplate} {runtimeComments}".Trim();
                    var filledTemplateResult = namedTests.FillAllParameters(finalReason, parameters);
                    if (!string.IsNullOrWhiteSpace(filledTemplateResult.ProcessedText))
                    {
                        // Add as a list item
                        failureMessages.AppendLine($"<li>{filledTemplateResult.ProcessedText}</li>");
                    }
                }
            }

            // Recurse into all sub-tests to find other failures.
            foreach (var subResult in testResult.SubTestResults)
            {
                BuildRejectionBody(subResult, failureMessages, namedTests, parameters);
            }
        }


        private async Task SendOutlookEmailAsync(OutgoingEmail emailData, Microsoft.Office.Interop.Outlook.Application outlookApp)
        {
            Microsoft.Office.Interop.Outlook.MailItem mailItem = null;
            try
            {
                mailItem = outlookApp.CreateItem(Microsoft.Office.Interop.Outlook.OlItemType.olMailItem);

                mailItem.Subject = emailData.Subject;
                mailItem.HTMLBody = emailData.Body;

                foreach (var recipient in emailData.ToRecipients)
                {
                    mailItem.Recipients.Add(recipient).Type = (int)Microsoft.Office.Interop.Outlook.OlMailRecipientType.olTo;
                }
                foreach (var recipient in emailData.CcRecipients)
                {
                    mailItem.Recipients.Add(recipient).Type = (int)Microsoft.Office.Interop.Outlook.OlMailRecipientType.olCC;
                }
                foreach (var recipient in emailData.BccRecipients)
                {
                    mailItem.Recipients.Add(recipient).Type = (int)Microsoft.Office.Interop.Outlook.OlMailRecipientType.olBCC;
                }

                if (!mailItem.Recipients.ResolveAll())
                {
                    Log.RecordError("Could not resolve all email recipients.", null, nameof(SendOutlookEmailAsync));
                }

                // Add Attachments
                foreach (var attachmentPath in emailData.Attachments)
                {
                    if (File.Exists(attachmentPath))
                    {
                        mailItem.Attachments.Add(attachmentPath);
                    }
                }

                mailItem.Send();

                Log.RecordMessage("Confirmation email sent successfully.", BisLogMessageType.Note);
            }
            catch (Exception ex)
            {
                Log.RecordError("Failed to send Outlook email.", ex, nameof(SendOutlookEmailAsync));
                ArcGIS.Desktop.Framework.Dialogs.MessageBox.Show("Failed to send confirmation email. Please check the logs.", "Email Error");
            }
            finally
            {
                if (mailItem != null) Marshal.ReleaseComObject(mailItem);
            }
        }
    }
}