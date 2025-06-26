using IC_Loader_Pro.Models;
using IC_Rules_2025;
using System;
using System.Text.RegularExpressions;
using static IC_Loader_Pro.Module1;
using static Bis_Regex;

namespace IC_Loader_Pro.Services
{
    public class EmailClassifierService
    {
        private readonly IC_Rules _rulesEngine;
        private readonly Bis_Regex bis_Regex;

        public EmailClassifierService()
        {
            // Get the singleton instance of the rules engine from Module1
            _rulesEngine = IcRules;
            if (_rulesEngine == null)
            {
                throw new InvalidOperationException("IC_Rules engine has not been initialized in Module1.");
            }

            bis_Regex = new Bis_Regex(Log);
        }

        /// <summary>
        /// Determines the type of an email based on its subject line and content.
        /// Replicates the functionality of the legacy determineEmailType method.
        /// </summary>
        /// <param name="email">The EmailItem to classify.</param>
        /// <returns>An EmailClassificationResult object.</returns>
        public EmailClassificationResult ClassifyEmail(EmailItem email)
        {
            var result = new EmailClassificationResult();

            // Rule 1: Check for empty subject
            if (string.IsNullOrWhiteSpace(email.Subject))
            {
                result.Type = EmailType.EmptySubjectline;
                result.IsSubjectLineValid = false;
                result.InvalidReason = "Email subject is empty.";
                return result;
            }

            // Note: The named regex patterns like "SubjectLineIsSpam", "SubjectLineIsCEA", etc.,
            // must be configured in your BIS_Tools_Core.BIS_Regex class and its underlying data source.

            // Rule 2: Check for Spam
            if (bis_Regex.StringMatchesNamedRegex("SubjectLineIsSpam", email.Subject) ||
                bis_Regex.StringMatchesNamedRegex("SenderIsSpam", email.SenderEmailAddress))
            {
                result.Type = EmailType.Spam;
                return result;
            }

            // Rule 3: Check for Auto-Reply
            if (bis_Regex.StringMatchesNamedRegex("autoReplyEmail", email.Subject))
            {
                result.Type = EmailType.AutoResponse;
                return result;
            }

            // Rule 4: Check for specific IC types based on subject line
            if (bis_Regex.StringMatchesNamedRegex("SubjectLineIsCEA", email.Subject))
            {
                result.Type = EmailType.CEA;
                return result;
            }
            if (bis_Regex.StringMatchesNamedRegex("SubjectLineIsDNA", email.Subject))
            {
                result.Type = EmailType.DNA;
                return result;
            }
            if (bis_Regex.StringMatchesNamedRegex("SubjectLineIsCKE", email.Subject))
            {
                result.Type = EmailType.CKE;
                return result;
            }
            if (bis_Regex.StringMatchesNamedRegex("SubjectLineIsIEC", email.Subject))
            {
                result.Type = EmailType.IEC;
                return result;
            }
            if (bis_Regex.StringMatchesNamedRegex("SubjectLineIsWRS", email.Subject))
            {
                result.Type = EmailType.WRS;
                return result;
            }

            // Add other rules (EDD_Resubmit, etc.) here as needed.

            // If no other rule matches, the type remains "Unknown"
            return result;
        }
    }
}