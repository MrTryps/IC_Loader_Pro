using BIS_Tools_DataModels_2025;
using IC_Loader_Pro.Models;
using IC_Rules_2025;
using System;
using System.Linq;

namespace IC_Loader_Pro.Services
{
    public class EmailClassifierService
    {
        private readonly IC_Rules _rulesEngine;
        private readonly BIS_Log _log;

        public EmailClassifierService(IC_Rules rulesEngine, BIS_Log log)
        {
            _rulesEngine = rulesEngine ?? throw new ArgumentNullException(nameof(rulesEngine));
            _log = log ?? throw new ArgumentNullException(nameof(log));
        }

        public EmailClassificationResult ClassifyEmail(EmailItem email)
        {
            const string methodName = "ClassifyEmail";
            var result = new EmailClassificationResult();
            var regexTool = Module1.RegexTool;

            if (regexTool == null)
            {
                _log.RecordError("The RegexTool within IC_Rules is not initialized.", null, methodName);
                throw new InvalidOperationException("RegexTool is not available.");
            }

            try
            {
                // --- Start of New ID Extraction Logic ---

                // 1. Clean and split the subject line into unique words
                string cleanedSubject = email.Subject ?? string.Empty;
                // NOTE: The VB code calls a 'cleanSubjectLine' method which is not provided.
                // We will replicate the splitting logic. You may need to add more cleaning steps later.
                var delimiters = new[] { ',', ';', '/', '&', '(', ')', '_', '-' };
                var subjectWords = cleanedSubject.Split(delimiters, StringSplitOptions.RemoveEmptyEntries)
                                                 .Select(w => w.Trim().ToUpper())
                                                 .Where(w => !string.IsNullOrEmpty(w))
                                                 .Distinct()
                                                 .ToList();

                // 2. Iterate through each word and check it against the regex rules
                foreach (var word in subjectWords)
                {
                    // Check for Preference IDs
                    result.PrefIds.AddRange(regexTool.ReturnMatchesOfNamedRegex("WordIsPrefId", word));

                    // Check for SRP IDs (AltIds)
                    result.AltIds.AddRange(regexTool.ReturnMatchesOfNamedRegex("WordIsSrpId", word));

                    // Check for Activity Numbers
                    result.ActivityNums.AddRange(regexTool.ReturnMatchesOfNamedRegex("WordIsActivityNum", word));
                }

                // --- End of New ID Extraction Logic ---


                // Now, perform the rest of the classification (spam, auto-reply, etc.)
                if (string.IsNullOrWhiteSpace(email.Subject))
                {
                    result.Type = EmailType.EmptySubjectline;
                    return result;
                }

                if (regexTool.StringMatchesNamedRegex("SubjectLineIsSpam", email.Subject) ||
                    regexTool.StringMatchesNamedRegex("SenderIsSpam", email.SenderEmailAddress))
                {
                    result.Type = EmailType.Spam;
                    return result;
                }

                if (regexTool.StringMatchesNamedRegex("autoReplyEmail", email.Subject))
                {
                    result.Type = EmailType.AutoResponse;
                    return result;
                }

                if (regexTool.StringMatchesNamedRegex("SubjectLineIsCEA", email.Subject))
                {
                    result.Type = EmailType.CEA;
                }
                
                if (regexTool.StringMatchesNamedRegex("SubjectLineIsCKE", email.Subject))
                {
                    result.Type = EmailType.CKE;
                }

                if (regexTool.StringMatchesNamedRegex("SubjectLineIsDN", email.Subject))
                {
                    result.Type = EmailType.DNA;
                }

                if (regexTool.StringMatchesNamedRegex("SubjectLineIsIEC", email.Subject))
                {
                    result.Type = EmailType.IEC;
                }

                if (regexTool.StringMatchesNamedRegex("SubjectLineIsWRS", email.Subject))
                {
                    result.Type = EmailType.WRS;
                }
            }
            catch (Exception ex)
            {
                _log.RecordError($"An unexpected error occurred during email classification for subject: '{email.Subject}'", ex, methodName);
                // We return the partial result instead of throwing, so the app can handle it gracefully.
            }

            return result;
        }
    }
}