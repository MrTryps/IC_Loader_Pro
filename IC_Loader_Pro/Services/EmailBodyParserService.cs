// In IC_Loader_Pro/Services/EmailBodyParserService.cs

using IC_Rules_2025;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text.RegularExpressions;
using static IC_Loader_Pro.Module1;

namespace IC_Loader_Pro.Services
{
    /// <summary>
    /// A service to parse and extract structured data from the body of an email
    /// based on a set of rules defined in the database.
    /// </summary>
    public class EmailBodyParserService
    {
        private readonly List<BodyFieldRule> _fieldRules = new List<BodyFieldRule>();
        private readonly string _icType;

        // Represents a single rule for finding a field in the email body
        private class BodyFieldRule
        {
            public string FieldName { get; }
            public string StringToSearch { get; }
            public string ValueRule { get; }

            public BodyFieldRule(string fieldName, string stringToSearch, string valueRule)
            {
                FieldName = fieldName;
                StringToSearch = stringToSearch;
                ValueRule = valueRule;
            }
        }

        public EmailBodyParserService(string icType)
        {
            _icType = icType;
            LoadFieldRules();
        }

        private void LoadFieldRules()
        {
            var paramDict = new Dictionary<string, object> { { "IcType", _icType } };
            var dt = PostGreTool.ExecuteNamedQuery("IncomingEmailBodyFieldRules", paramDict) as DataTable;

            if (dt == null) return;

            foreach (DataRow dr in dt.Rows)
            {
                string fieldName = dr["FIELDNAME"]?.ToString();
                string stringToSearch = dr["STRINGTOSEARCH"]?.ToString();
                string valueRule = dr["VALUERULE"]?.ToString();
                _fieldRules.Add(new BodyFieldRule(fieldName, stringToSearch, valueRule));
            }
        }

        /// <summary>
        /// Main method to parse the body text and extract key-value pairs.
        /// </summary>
        public Dictionary<string, string> GetFieldsFromBody(string bodyText)
        {
            var dataFound = new Dictionary<string, string>();
            if (string.IsNullOrWhiteSpace(bodyText)) return dataFound;

            var partitionedBody = BreakBodyTextByFieldNames(bodyText);

            foreach (string line in partitionedBody)
            {
                var (found, fieldName, fieldValue) = CheckLineForField(line);
                if (found)
                {
                    dataFound[fieldName] = fieldValue;
                }
            }
            return dataFound;
        }

        private List<string> BreakBodyTextByFieldNames(string bodyText)
        {
            const string lineBreakSymbol = "~~~~";

            // Clean the body text similar to the VB logic
            string cleanedBodyText = Regex.Replace(bodyText, @"[\u0000-\u001F\u007F-\u00FF]", "");
            cleanedBodyText = cleanedBodyText.Replace("\"", "");
            cleanedBodyText = Regex.Replace(cleanedBodyText, @"\s+", " ").Trim();

            if (string.IsNullOrEmpty(cleanedBodyText)) return new List<string>();

            // Insert break symbols before each field name found
            foreach (var rule in _fieldRules)
            {
                cleanedBodyText = Regex.Replace(cleanedBodyText,
                                                Regex.Escape(rule.StringToSearch),
                                                lineBreakSymbol + rule.StringToSearch,
                                                RegexOptions.IgnoreCase);
            }

            return cleanedBodyText.Split(new[] { lineBreakSymbol }, System.StringSplitOptions.RemoveEmptyEntries).ToList();
        }

        private (bool, string, string) CheckLineForField(string line)
        {
            if (string.IsNullOrWhiteSpace(line)) return (false, null, null);

            foreach (var rule in _fieldRules)
            {
                // Check if the line starts with the string we're looking for
                if (line.Trim().StartsWith(rule.StringToSearch, System.StringComparison.OrdinalIgnoreCase))
                {
                    // The value is the part of the line after the search string
                    string value = line.Substring(rule.StringToSearch.Length).Trim();

                    // Simple value rule: if the value is empty, use the rule's default
                    if (string.IsNullOrWhiteSpace(value))
                    {
                        value = rule.ValueRule;
                    }

                    return (true, rule.FieldName, value);
                }
            }
            return (false, null, null);
        }
    }
}