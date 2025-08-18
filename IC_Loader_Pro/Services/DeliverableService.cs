using System;
using System.Collections.Generic;
using System.Threading.Tasks;
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
        /// <returns>The new unique Deliverable ID as a string.</returns>
        public async Task<string> CreateNewDeliverableRecordAsync(string deliveryMethod)
        {
            const string methodName = "CreateNewDeliverableRecordAsync";
            string userName = Environment.UserName;

            // Build the dictionary of parameters for the named query
            var paramDict = new Dictionary<string, object>
            {
                { "DELIVERYMETHOD", deliveryMethod.ToUpper() },
                { "LogPath", Log.FileNameAndPath},
                { "USERNAME", userName }
            };

            string newDeliverableId;
            try
            {
                // Run the database query on a background thread
                newDeliverableId = await Task.Run(() =>
                    PostGreTool.ExecuteNamedQuery("newGisDeliverable", paramDict) as string
                );
            }
            catch (Exception ex)
            {
                Log.RecordError("Error creating new GIS Deliverable record in the database.", ex, methodName);
                // Re-throw the exception so the calling code knows the operation failed
                throw;
            }

            if (string.IsNullOrEmpty(newDeliverableId))
            {
                throw new Exception("The database did not return a new Deliverable ID after creating the record.");
            }

            return newDeliverableId;
        }       
    }
}