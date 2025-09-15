using BIS_Tools_DataModels_2025;
using IC_Loader_Pro.Models;
using IC_Rules_2025;
using System;
using System.Collections.Generic;
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
                // --- Pre-Checks (Unchanged) ---
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

                // --- START OF UNIFIED LOGIC ---
                // 1. Split the subject line into words just ONCE.
                string cleanedSubject = email.Subject ?? string.Empty;
                var delimiters = new[] { ' ', ',', ';', '/', '&', '(', ')', '_', '-' };
                var subjectWords = cleanedSubject.Split(delimiters, StringSplitOptions.RemoveEmptyEntries)
                                                 .Select(w => w.Trim().ToUpper())
                                                 .Where(w => !string.IsNullOrEmpty(w))
                                                 .Distinct()
                                                 .ToList();

                var foundTypes = new HashSet<EmailType>(); // Use a HashSet to avoid duplicate types

                // 2. Iterate through each word and check it against ALL relevant rules.
                foreach (var word in subjectWords)
                {
                    // ID Extraction
                    result.PrefIds.AddRange(regexTool.ReturnMatchesOfNamedRegex("WordIsPrefId", word));
                    result.AltIds.AddRange(regexTool.ReturnMatchesOfNamedRegex("WordIsSrpId", word));
                    result.ActivityNums.AddRange(regexTool.ReturnMatchesOfNamedRegex("WordIsActivityNum", word));

                    // Email Type Classification
                    if (regexTool.StringMatchesNamedRegex("SubjectLineIsCEA", word)) foundTypes.Add(EmailType.CEA);
                    if (regexTool.StringMatchesNamedRegex("SubjectLineIsCKE", word)) foundTypes.Add(EmailType.CKE);
                    if (regexTool.StringMatchesNamedRegex("SubjectLineIsDN", word)) foundTypes.Add(EmailType.DNA);
                    if (regexTool.StringMatchesNamedRegex("SubjectLineIsIEC", word)) foundTypes.Add(EmailType.IEC);
                    if (regexTool.StringMatchesNamedRegex("SubjectLineIsWRS", word)) foundTypes.Add(EmailType.WRS);
                }

                // 3. Now, determine the final type based on how many unique matches we found.
                if (foundTypes.Count == 1)
                {
                    result.Type = foundTypes.First();
                }
                else if (foundTypes.Count > 1)
                {
                    result.Type = EmailType.Multiple;
                }
                else
                {
                    result.Type = EmailType.Unknown;
                }
                // --- END OF UNIFIED LOGIC ---
            }
            catch (Exception ex)
            {
                _log.RecordError($"An unexpected error occurred during email classification for subject: '{email.Subject}'", ex, methodName);
            }

            return result;
        }
    }
}