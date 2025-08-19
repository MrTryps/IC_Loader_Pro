using IC_Loader_Pro.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using static BIS_Log;
using static IC_Loader_Pro.Module1; // Provides static access to Log and PostGreTool

namespace IC_Loader_Pro.Services
{
    /// <summary>
    /// A service dedicated to interacting with the deliverables table in the database.
    /// </summary>
    public class DeliverableService
    {

        /// <summary>
        /// Creates a new deliverable record in the database and returns the new ID.
        /// </summary>
        /// <param name="deliveryMethod">The method of delivery (e.g., "EMAIL", "MANUAL").</param>
        /// <param name="icType">The IC Type of the submission (e.g., "CEA", "DNA").</param>
        /// <param name="prefId">The Preference ID for the submission.</param>
        /// <param name="processDate">The date of the submission (e.g., email received time).</param>
        /// <returns>The new unique Deliverable ID as a string.</returns>
        public async Task<string> CreateNewDeliverableRecordAsync(string deliveryMethod, string icType, string prefId, DateTime processDate)
        {
            const string methodName = "CreateNewDeliverableRecordAsync";
            string userName = Environment.UserName;

            // Build the dictionary of parameters to match the updated database function
            var paramDict = new Dictionary<string, object>
    {
        { "deliverymethod", deliveryMethod.ToUpper() },
        { "logpath", Log.FileNameAndPath },
        { "username", userName },
        { "ic_type_in", icType },
        { "pref_id_in", prefId },
        { "process_date_in", processDate }
    };

            string newDeliverableId;
            try
            {
                newDeliverableId = await Task.Run(() =>
                    PostGreTool.ExecuteNamedQuery("newGisDeliverable", paramDict) as string
                );
            }
            catch (Exception ex)
            {
                Log.RecordError("Error creating new GIS Deliverable record in the database.", ex, methodName);
                throw;
            }

            if (string.IsNullOrEmpty(newDeliverableId))
            {
                throw new Exception("The database did not return a new Deliverable ID after creating the record.");
            }

            return newDeliverableId;
        }

        /// <summary>
        /// Updates the srp_gis_deliverable_emailinfo table with details from the processed email.
        /// </summary>
        public async Task UpdateEmailInfoRecordAsync(string deliverableId, EmailItem email, EmailClassificationResult classification, string sourceFolder)
        {
            const string methodName = "UpdateEmailInfoRecordAsync";

            var setClauses = new List<string>();
            // Use a List<object> to ensure the parameter order is correct for ODBC
            var parameters = new List<object>();

            // Dynamically build the SET clauses, using '?' as the placeholder
            if (!string.IsNullOrEmpty(email.Subject))
            {
                setClauses.Add("subjectline = ?");
                parameters.Add(email.Subject);
            }
            if (email.ReceivedTime > DateTime.MinValue)
            {
                setClauses.Add("senddate = ?");
                parameters.Add(email.ReceivedTime);
            }
            if (!string.IsNullOrEmpty(email.Emailid))
            {
                setClauses.Add("emailid = ?");
                parameters.Add(email.Emailid);
            }
            if (classification.PrefIds.Any())
            {
                setClauses.Add("subjectlineprefids = ?");
                parameters.Add(string.Join(";", classification.PrefIds));
            }
            if (!string.IsNullOrEmpty(sourceFolder))
            {
                setClauses.Add("outlookfolder = ?");
                parameters.Add(sourceFolder);
            }
            if (!string.IsNullOrEmpty(classification.Note))
            {
                setClauses.Add("notes = ?");
                parameters.Add(classification.Note);
            }

            if (!setClauses.Any())
            {
                Log.RecordMessage("No email information to update in the database.", BisLogMessageType.Note);
                return;
            }

            // The parameter for the WHERE clause must be the LAST one added to the list
            string sql = $"UPDATE srp_gis_deliverable_emailinfo SET {string.Join(", ", setClauses)} WHERE deliverable_id = ?";
            parameters.Add(deliverableId);

            try
            {
                // Call ExecuteRawQuery, which is designed to work with '?' placeholders
                await Task.Run(() => PostGreTool.ExecuteRawQuery(sql, parameters, "NOTHING"));
            }
            catch (Exception ex)
            {
                Log.RecordError($"Error updating email info record for Deliverable ID '{deliverableId}'.", ex, methodName);
                throw;
            }
        }

        /// <summary>
        /// Updates the srp_gis_contactinfo table with the submitter's contact details.
        /// </summary>
        public async Task UpdateContactInfoRecordAsync(string deliverableId, EmailItem email)
        {
            const string methodName = "UpdateContactInfoRecordAsync";

            var setClauses = new List<string>();
            var parameters = new List<object>();

            // Dynamically build the SET clauses based on available data
            if (!string.IsNullOrEmpty(email.SenderName))
            {
                setClauses.Add("submitter_name = ?");
                parameters.Add(email.SenderName);
            }
            if (!string.IsNullOrEmpty(email.SenderEmailAddress))
            {
                setClauses.Add("submitter_email = ?");
                parameters.Add(email.SenderEmailAddress);
            }

            if (!setClauses.Any())
            {
                Log.RecordMessage("No contact information to update in the database.", BisLogMessageType.Note);
                return;
            }

            // Finalize the SQL statement
            string sql = $"UPDATE srp_gis_contactinfo SET {string.Join(", ", setClauses)} WHERE deliverable_id = ?";
            parameters.Add(deliverableId);

            try
            {
                await Task.Run(() => PostGreTool.ExecuteRawQuery(sql, parameters, "NOTHING"));
            }
            catch (Exception ex)
            {
                Log.RecordError($"Error updating contact info record for Deliverable ID '{deliverableId}'.", ex, methodName);
                throw;
            }
        }

        // --- REPLACE the existing UpdateBodyDataRecordAsync method with this one ---
        /// <summary>
        /// Updates the srp_gis_body_data table with data parsed from the email body.
        /// </summary>
        public async Task UpdateBodyDataRecordAsync(string deliverableId, Dictionary<string, string> bodyData)
        {
            const string methodName = "UpdateBodyDataRecordAsync";

            if (bodyData == null || !bodyData.Any())
            {
                Log.RecordMessage("No parsed body data to update in the database.", BisLogMessageType.Note);
                return;
            }

            // A dictionary mapping column names to their maximum allowed length.
            var columnLengths = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase)
    {
        { "gisnameaddress", 255 },
        { "lsrpnameid", 50 },
        { "lsrpemail", 120 },
        { "gisprofname", 100 },
        { "gisprofemail", 120 },
        { "gisprofphone", 20 },
        { "prefid", 20 },
        { "sitename", 120 },
        { "siteaddress", 200 },
        { "filesuffix", 10 },
        { "subjectitemid", 20 },
        { "remedialaction", 50 },
        { "siteboundary", 50 }
    };

            var setClauses = new List<string>();
            var parameters = new List<object>();

            foreach (var kvp in bodyData)
            {
                string columnName = kvp.Key;
                string value = kvp.Value;

                // Check if the column is a valid, known column
                if (columnLengths.ContainsKey(columnName))
                {
                    int maxLength = columnLengths[columnName];

                    // Check if the value is too long
                    if (value.Length > maxLength)
                    {
                        Log.RecordMessage($"Data for column '{columnName}' is too long ({value.Length} chars). Truncating to {maxLength} chars.", BisLogMessageType.Warning);
                        // Truncate the value before adding it
                        value = value.Substring(0, maxLength);
                    }

                    setClauses.Add($"{columnName} = ?");
                    parameters.Add(value);
                }
                else
                {
                    Log.RecordMessage($"Ignored unknown field '{columnName}' from email body parser.", BisLogMessageType.Warning);
                }
            }

            if (!setClauses.Any())
            {
                Log.RecordMessage("No valid body data fields were found to update.", BisLogMessageType.Note);
                return;
            }

            string sql = $"UPDATE srp_gis_body_data SET {string.Join(", ", setClauses)} WHERE deliverable_id = ?";
            parameters.Add(deliverableId);

            try
            {
                await Task.Run(() => PostGreTool.ExecuteRawQuery(sql, parameters, "NOTHING"));
            }
            catch (Exception ex)
            {
                Log.RecordError($"Error updating body data record for Deliverable ID '{deliverableId}'.", ex, methodName);
                throw;
            }
        }

    }


}



