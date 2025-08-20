using ArcGIS.Core.Data;
using ArcGIS.Core.Geometry;
using ArcGIS.Desktop.Framework.Threading.Tasks;
using IC_Loader_Pro.Models;
using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using static BIS_Log;
using static IC_Loader_Pro.Module1;
using Exception = System.Exception;

namespace IC_Loader_Pro.Services
{
    public class ShapeProcessingService
    {
        /// <summary>
        /// Gets the next available Shape ID from the database by calling a named query.
        /// </summary>
        /// <param name="deliverableId">The parent Deliverable ID for this shape.</param>
        /// <param name="idPrefix">The IC Type specific prefix for the new ID (e.g., "CEA", "DNA").</param>
        /// <returns>A new, unique shape ID as a string.</returns>
        public async Task<string> GetNextShapeIdAsync(string deliverableId, string idPrefix)
        {
            const string methodName = "GetNextShapeIdAsync";

            if (string.IsNullOrEmpty(deliverableId))
                throw new ArgumentNullException(nameof(deliverableId));
            if (string.IsNullOrEmpty(idPrefix))
                throw new ArgumentNullException(nameof(idPrefix));

            var paramDict = new Dictionary<string, object>
    {
        { "GIS_ID", deliverableId },
        { "IC_PREFIX", idPrefix }
    };

            string newShapeId;
            try
            {
                // Run the database query on a background thread
                newShapeId = await Task.Run(() =>
                    PostGreTool.ExecuteNamedQuery("RETURN_NEXT_SHAPE_ID", paramDict) as string
                );
            }
            catch (Exception ex)
            {
                Log.RecordError("Error obtaining new shape ID from the database.", ex, methodName);
                throw; // Re-throw the exception so the calling OnSave command can handle it
            }

            if (string.IsNullOrEmpty(newShapeId))
            {
                throw new Exception("Database did not return a new shape ID.");
            }

            return newShapeId;
        }

        /// <summary>
        /// Creates the initial record for a shape in the shape_info table.
        /// </summary>
        public async Task RecordShapeInfoAsync(string newShapeId, string submissionId, string deliverableId, string prefId, string icType)
        {
            const string methodName = "RecordShapeInfoAsync";

            // 1. Get the name of the shape info table from the rules engine
            var shapeInfoTableRule = IcRules.ReturnIcGisTypeSettings(icType)?.ShapeInfoTable;
            if (shapeInfoTableRule == null || string.IsNullOrEmpty(shapeInfoTableRule.PostGreFeatureClassName))
            {
                throw new InvalidOperationException($"The 'shape_info_table' is not configured for IC Type '{icType}'.");
            }
            string tableName = shapeInfoTableRule.PostGreFeatureClassName;

            // 2. Build the dynamic INSERT statement and parameter list
            var columns = new Dictionary<string, object>
    {
        { "SHAPE_ID", newShapeId },
        { "submission_id", submissionId },
        { "deliverable_id", deliverableId },
        { "pref_id", prefId },
        { "ic_type", icType }
    };

            var fieldList = string.Join(", ", columns.Keys);
            var valuePlaceholders = string.Join(", ", columns.Keys.Select(k => "?"));
            var sql = $"INSERT INTO {tableName} ({fieldList}) VALUES ({valuePlaceholders})";

            var parameters = columns.Values.ToList();

            // 3. Execute the query
            try
            {
                await Task.Run(() => PostGreTool.ExecuteRawQuery(sql, parameters, "NOTHING"));
                Log.RecordMessage($"Created record in {tableName} for Shape ID: {newShapeId}", BisLogMessageType.Note);
            }
            catch (Exception ex)
            {
                Log.RecordError($"Error creating new shape info record in '{tableName}' for Shape ID '{newShapeId}'.", ex, methodName);
                throw;
            }
        }

        /// <summary>
        /// Checks for duplicate polygons in the proposed feature class by comparing both
        /// geometry (with a tolerance) and key business attributes.
        /// </summary>
        public async Task<bool> IsDuplicateInProposedAsync(Geometry shape, string currentPrefId, string icType)
        {
            bool isDuplicate = false;
            await QueuedTask.Run(async () =>
            {
                // 1. Get the 'proposed' feature class from the rules
                var proposedFcRule = IcRules.ReturnIcGisTypeSettings(icType)?.ProposedFeatureClass;
                if (proposedFcRule == null || string.IsNullOrEmpty(proposedFcRule.PostGreFeatureClassName))
                {
                    Log.RecordError("Proposed feature class is not configured.", null, nameof(IsDuplicateInProposedAsync));
                    return; // Cannot check for duplicates if the layer isn't configured
                }

                // This is a placeholder for opening the actual feature class
                // In a real implementation, you would use the workspace info from proposedFcRule
                // to connect to the geodatabase and open the feature class.
                // For now, we'll assume it can be opened directly for shelling purposes.
                // FeatureClass proposedFc = ... open the feature class ...
                // if (proposedFc == null) return;

                // 2. Create a spatial filter to find intersecting features efficiently
                var spatialFilter = new SpatialQueryFilter
                {
                    FilterGeometry = shape.Extent,
                    SpatialRelationship = SpatialRelationship.Intersects
                };

                // 3. (Shelled) This block simulates searching the feature class
                // In a real implementation, you would use: using (var cursor = proposedFc.Search(spatialFilter, false))
                Log.RecordMessage("Searching for spatially similar shapes in proposed layer (shelled)...", BisLogMessageType.Note);

                // --- Placeholder: Simulate finding one potential match ---
                var potentialMatches = new List<string> { "CEA_12345" }; // A fake existing shape ID
                                                                         // ---

                foreach (var existingShapeId in potentialMatches) // In reality: loop through cursor results
                {
                    // In reality: Get the geometry of the existing feature from the cursor
                    // Geometry existingShape = cursor.Current.GetShape();

                    // 4. (Shelled) Use GeometryEngine to check for equality with tolerance
                    // bool geometriesAreEqual = GeometryEngine.Instance.Equals(shape, existingShape);
                    bool geometriesAreEqual = true; // Assume they are equal for this shelled version

                    if (geometriesAreEqual)
                    {
                        // 5. If geometries match, check the business rules via the shape_info table
                        var existingShapeInfo = await GetShapeInfoAsync(existingShapeId);

                        if (existingShapeInfo != null &&
                            existingShapeInfo.PrefId.Equals(currentPrefId, StringComparison.OrdinalIgnoreCase) &&
                            existingShapeInfo.IcType.Equals(icType, StringComparison.OrdinalIgnoreCase) &&
                            (existingShapeInfo.Status == "To Be Reviewed" || existingShapeInfo.Status == "Shape Approved"))
                        {
                            Log.RecordMessage($"Found a valid duplicate. New shape is a duplicate of existing shape ID: {existingShapeId}", BisLogMessageType.Note);
                            isDuplicate = true;
                            return; // Exit the loop as soon as the first valid duplicate is found
                        }
                    }
                }
            });

            return isDuplicate;
        }

        /// <summary>
        /// Updates a specific field for a record in the shape_info table.
        /// </summary>
        public async Task UpdateShapeInfoFieldAsync(string shapeId, string fieldName, object value)
        {
            Log.RecordMessage($"Updating {fieldName} for Shape ID: {shapeId} (shelled)...", BisLogMessageType.Note);
            // TODO: Create a named query that takes shapeId, fieldName, and value to update the table.
            await Task.CompletedTask;
        }

        /// <summary>
        /// (Shelled) Copies an approved shape's geometry into the proposed feature class.
        /// </summary>
        public async Task CopyShapeToProposedAsync(Geometry shape, string newShapeId)
        {
            Log.RecordMessage($"Copying shape {newShapeId} to proposed feature class (shelled)...", BisLogMessageType.Note);
            // TODO: Implement the GIS logic to create a new feature in the proposed feature class.
            await Task.CompletedTask;
        }

        /// <summary>
        /// (Shelled) A helper method to retrieve a record from the shape_info table.
        /// </summary>
        private async Task<(string PrefId, string IcType, string Status)?> GetShapeInfoAsync(string shapeId)
        {
            Log.RecordMessage($"Querying shape_info table for Shape ID: {shapeId} (shelled)...", BisLogMessageType.Note);
            // TODO: Create a named query "returnShapeInfo" that takes a shape_id
            // and returns the pref_id, ic_type, and shape_status fields.

            await Task.CompletedTask;

            // Return a sample record for testing purposes
            return (PrefId: "G000033577", IcType: "CEA", Status: "To Be Reviewed");
        }
    }
}
