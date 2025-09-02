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
        public async Task<bool> SendConfirmationEmailAsync(string deliverableId, IcTestResult testResult, string icType, Microsoft.Office.Interop.Outlook.Application outlookApp)
        {
            // 1. Build the initial email object.
            var emailToEdit = BuildReplyEmail(deliverableId, testResult, icType);
            if (emailToEdit == null)
            {
                Log.RecordMessage("Email not built because no recipients were specified.", BisLogMessageType.Note);
                return true;
            }

            bool userClickedSend = false;
            // 2. Show the preview window on the UI thread.
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

            // 3. Only send the email if the user clicked the "Send" button.
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


        private OutgoingEmail BuildReplyEmail(string deliverableId, IcTestResult testResult, string icType)
        {
            var finalCumulativeAction = testResult.CumulativeAction;

            if (finalCumulativeAction.EmailRp != true && finalCumulativeAction.EmailHazsite != true)
            {
                return null;
            }

            var outgoingEmail = new OutgoingEmail();
            var deliverableInfo = IcRules.ReturnEmailDeliverableInfo(deliverableId);

            // Set Recipients
            if (finalCumulativeAction.EmailRp == true && !string.IsNullOrEmpty(deliverableInfo.SenderEmail))
            {
                outgoingEmail.ToRecipients.Add(deliverableInfo.SenderEmail);
              //  Log.RecordMessage("--> Decision: Added recipient.", BisLogMessageType.Note);
            }
            if (finalCumulativeAction.EmailHazsite == true)
            {
                outgoingEmail.BccRecipients.Add("SRPGIS@dep.nj.gov");
            }
            else
            {
              //  Log.RecordMessage("--> Decision: Did NOT add recipient because SenderEmail was empty.", BisLogMessageType.Warning);
            }
            if (!outgoingEmail.ToRecipients.Any() && !outgoingEmail.BccRecipients.Any())
            {
                return null;
            }

            // 1. Create an instance of the IcNamedTests class, which now contains our template-filling logic.
            var namedTests = new IcNamedTests(Log, PostGreTool);

            // 2. Prepare the dictionary of dynamic values. Note the keys do NOT have brackets.
            var parameters = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
    {
        { "DELID", deliverableId },
        { "PREFID", testResult.OutputParams.GetValueOrDefault("prefid", "N/A") },
        { "SUBJECTLINE", deliverableInfo.SubjectLine },
        { "SENDDATE", deliverableInfo.SubmitDate },
        {"Ic_Type", icType  }
    };

            // 3. Determine which templates to use.
            var rootRule = testResult.TestRule;
            string subjectTemplate;

            // 1. Assemble the complete template for the opening text into a single string.
            string openingTextTemplate = "<P><div align='right'>{Template_GIS_OutgoingEmailID_Small}</DIV></P>";

            // 2. Append the correct pass/fail message template.
            if (testResult.CumulativeAction.ResultAction == TestActionResponse.Fail)
            {
                subjectTemplate = rootRule.FailSubject.ReplacementText;
                openingTextTemplate += rootRule.FailMessage.ReplacementText;
                openingTextTemplate += "{Template_GIS_Rejected}"; // Add postscript for rejection
            }
            else // Pass
            {
                subjectTemplate = rootRule.PassSubject.ReplacementText;
                openingTextTemplate += rootRule.PassMessage.ReplacementText;
            }

            // 3. Now, run the complete, assembled templates through the FillAllParameters method.
            var subjectResult = namedTests.FillAllParameters(subjectTemplate, parameters);
            var bodyResult = namedTests.FillAllParameters(openingTextTemplate, parameters);

            // 4. Assign the fully processed text to the email object.
            outgoingEmail.Subject = subjectResult.ProcessedText;
            outgoingEmail.AddToOpeningText(bodyResult.ProcessedText);

            var allMissingParams = subjectResult.MissingParameters.Union(bodyResult.MissingParameters).ToList();
            if (allMissingParams.Any())
            {
                // 3. If parameters are missing, show the temporary MessageBox for debugging.
                string missingParamsMessage = "The following parameters were found in the email templates but were not provided:\n\n- " +
                                              string.Join("\n- ", allMissingParams);

                // This needs to run on the UI thread.
                FrameworkApplication.Current.Dispatcher.Invoke(() =>
                {
                    ArcGIS.Desktop.Framework.Dialogs.MessageBox.Show(missingParamsMessage, "Missing Email Parameters");
                });
            }

            //// 6. Add any postscript text from the rule.
            //if (rootRule.PostscriptText != null)
            //{
            //    foreach (var postscript in rootRule.PostscriptText)
            //    {
            //        outgoingEmail.AddToClosingText(namedTests.FillAllParameters(postscript, parameters));
            //    }
            //}
            //// --- END OF UPDATED LOGIC ---

            return outgoingEmail;
        }

        /// <summary>
        /// Converts a hierarchical IcTestResult object into a formatted HTML string.
        /// </summary>
        /// <param name="testResult">The root test result to format.</param>
        /// <returns>An HTML string representing the test result tree.</returns>
        private string ConvertTestResultToHtml(IcTestResult testResult)
        {
            if (testResult == null) return string.Empty;

            var stringBuilder = new StringBuilder();
            // Start the HTML with an unordered list
            stringBuilder.AppendLine("<ul>");
            // Call the recursive helper to build out the list items
            BuildHtmlListItems(testResult, stringBuilder);
            stringBuilder.AppendLine("</ul>");

            return stringBuilder.ToString();
        }

        /// <summary>
        /// A recursive helper that traverses the test result hierarchy and builds HTML list items.
        /// </summary>
        private void BuildHtmlListItems(IcTestResult testResult, StringBuilder sb)
        {
            sb.AppendLine("<li>");

            // Use green for pass, red for fail
            string statusColor = testResult.Passed ? "green" : "red";
            string statusText = testResult.Passed ? "PASS" : "FAIL";

            // Main test result line
            sb.Append($"<span style='color:{statusColor}; font-weight:bold;'>({statusText})</span> ");
            sb.Append($"<strong>{testResult.TestRule.Name}</strong>");

            // Add comments if any exist
            if (testResult.Comments.Any())
            {
                sb.Append($": <em>{string.Join("; ", testResult.Comments)}</em>");
            }

            // If there are sub-tests, recurse
            if (testResult.SubTestResults.Any())
            {
                sb.AppendLine("<ul>");
                foreach (var subResult in testResult.SubTestResults)
                {
                    BuildHtmlListItems(subResult, sb);
                }
                sb.AppendLine("</ul>");
            }

            sb.AppendLine("</li>");
        }

        // A simple helper to replace key-value pairs in a string.
        private string ReplacePlaceholders(string template, Dictionary<string, string> values)
        {
            if (string.IsNullOrEmpty(template)) return "";

            foreach (var kvp in values)
            {
                template = template.Replace(kvp.Key, kvp.Value);
            }
            return template;
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

                // TODO: The logic for embedded images using MAPI properties can be added here later.

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