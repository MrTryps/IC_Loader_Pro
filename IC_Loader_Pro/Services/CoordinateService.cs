using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using static IC_Loader_Pro.Module1; // Provides static access to Log and PostGreTool

namespace IC_Loader_Pro.Services
{
    /// <summary>
    /// A service for handling database operations related to site coordinates.
    /// </summary>
    public class CoordinateService
    {
        /// <summary>
        /// Updates or inserts the coordinates for a given Preference ID in the PostgreSQL database.
        /// </summary>
        /// <param name="prefId">The Preference ID.</param>
        /// <param name="xCoord">The X coordinate.</param>
        /// <param name="yCoord">The Y coordinate.</param>
        /// <param name="coordSource">The source of the coordinates (e.g., "NJEMS").</param>
        public async Task UpdatePrefIdCoordinatesInPostgresAsync(string prefId, double xCoord, double yCoord, string coordSource)
        {
            const string methodName = "UpdatePrefIdCoordinatesInPostgresAsync";

            var paramDict = new Dictionary<string, object>
            {
                { "PREFID", prefId },
                { "X_COORD", xCoord },
                { "Y_COORD", yCoord },
                { "COORD_SOURCE", coordSource }
            };

            try
            {
                await Task.Run(() =>
                    PostGreTool.ExecuteNamedQuery("EnterPrefCoords", paramDict)
                );
            }
            catch (Exception ex)
            {
                Log.RecordError($"Error updating coordinates for Pref ID '{prefId}' in PostgreSQL.", ex, methodName);
                // We log the error but don't re-throw, as failing to update the cache
                // should not stop the main application workflow.
            }
        }
    }
}