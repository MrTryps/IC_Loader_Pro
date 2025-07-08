using IC_Loader_Pro.Models;
using IC_Rules_2025;
using System;

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

            // It's good practice to get a reference to the regex tool once.
            var regexTool = _rulesEngine.RegexTool;
            if (regexTool == null)
            {
                _log.RecordError("The RegexTool within IC_Rules is not initialized.", null, methodName);
                throw new InvalidOperationException("RegexTool is not available.");
            }

            try
            {
                if (string.IsNullOrWhiteSpace(email.Subject))
                {
                    result.Type = EmailType.EmptySubjectline;
                    return result;
                }

                // Correctly call the method on the regexTool instance.
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

                // ... (other classification rules using regexTool) ...
                if (regexTool.StringMatchesNamedRegex("SubjectLineIsCEA", email.Subject))
                {
                    result.Type = EmailType.CEA;
                }
                // ... etc. ...
            }
            catch (Exception ex)
            {
                _log.RecordError($"An unexpected error occurred during email classification for subject: '{email.Subject}'", ex, methodName);
                throw;
            }

            return result;
        }
    }
}