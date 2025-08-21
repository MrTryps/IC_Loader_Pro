using ArcGIS.Core.Data;
using ArcGIS.Core.Geometry;
using ArcGIS.Desktop.Editing;
using ArcGIS.Desktop.Framework.Threading.Tasks;
using IC_Loader_Pro.Models;
using IC_Rules_2025;
using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Threading.Tasks;
using static BIS_Log;
using static IC_Loader_Pro.Module1;
using Exception = System.Exception;

namespace IC_Loader_Pro.Services
{
    public class ShapeProcessingService
    {
        private readonly IC_Rules _rules;
        private readonly BIS_Log _log;

        public ShapeProcessingService(IC_Rules rules, BIS_Log log)
        {
            _rules = rules;
            _log = log;
        }

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

            // All GIS and database operations must run on the background thread.
            await QueuedTask.Run(async () =>
            {
                // 1. Get the rule for the 'proposed' feature class from the IC_Rules engine.
                var proposedFcRule = _rules.ReturnIcGisTypeSettings(icType)?.ProposedFeatureClass;
                if (proposedFcRule == null || string.IsNullOrEmpty(proposedFcRule.PostGreFeatureClassName))
                {
                    _log.RecordError("Proposed feature class is not configured in the rules.", null, nameof(IsDuplicateInProposedAsync));
                    return;
                }

                var workspaceRule = proposedFcRule.WorkSpaceRule;
                Geodatabase geodatabase = null;
                try
                {
                    // 2. Build the connection properties for the enterprise geodatabase.
                    var dbConnectionProperties = new DatabaseConnectionProperties(EnterpriseDatabaseType.PostgreSQL)
                    {
                        // Format for PostgreSQL instance is "sde:postgresql:<server>"
                        Instance = $"sde:postgresql:{workspaceRule.Server}",
                        Database = workspaceRule.Database,
                        User = workspaceRule.User,
                        Password = workspaceRule.Password,
                        Version = workspaceRule.Version
                    };

                    // 3. Connect to the geodatabase.
                    geodatabase = new Geodatabase(dbConnectionProperties);

                    // 4. Open the target feature class. The 'using' statement ensures it's properly closed.
                    using (var proposedFc = geodatabase.OpenDataset<FeatureClass>(proposedFcRule.PostGreFeatureClassName))
                    {
                        // 5. Create a spatial filter to efficiently find candidate features.
                        var spatialFilter = new SpatialQueryFilter
                        {
                            FilterGeometry = shape.Extent,
                            SpatialRelationship = SpatialRelationship.Intersects
                        };

                        // 6. Search the feature class for potential duplicates.
                        using (var cursor = proposedFc.Search(spatialFilter, false))
                        {
                            while (cursor.MoveNext())
                            {
                                using (var currentFeature = cursor.Current as Feature)
                                {
                                    if (currentFeature == null) continue;

                                    // 7. Use GeometryEngine to check for equality with tolerance.
                                    if (GeometryEngine.Instance.Equals(shape, currentFeature.GetShape()))
                                    {
                                        // 8. If geometries match, check the business rules via the shape_info table.
                                        string existingShapeId = currentFeature["SHAPE_ID"]?.ToString();
                                        if (string.IsNullOrEmpty(existingShapeId)) continue;

                                        var existingShapeInfo = await GetShapeInfoAsync(existingShapeId);

                                        if (existingShapeInfo.HasValue &&
                                            existingShapeInfo.Value.PrefId.Equals(currentPrefId, StringComparison.OrdinalIgnoreCase) &&
                                            existingShapeInfo.Value.IcType.Equals(icType, StringComparison.OrdinalIgnoreCase) &&
                                            (existingShapeInfo.Value.Status == "To Be Reviewed" || existingShapeInfo.Value.Status == "Shape Approved"))
                                        {
                                            _log.RecordMessage($"Found a valid duplicate. New shape is a duplicate of existing shape ID: {existingShapeId}", BisLogMessageType.Note);
                                            isDuplicate = true;
                                            return; // Exit as soon as the first valid duplicate is found
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    _log.RecordError($"Failed to connect to or query the proposed feature class '{proposedFcRule.PostGreFeatureClassName}'.", ex, nameof(IsDuplicateInProposedAsync));
                    return;
                }
                finally
                {
                    // Ensure the geodatabase connection is always closed and disposed of.
                    geodatabase?.Dispose();
                }
            });

            return isDuplicate;
        }


        /// <summary>
        /// Updates a single field for a given shape's record in the shape_info table.
        /// </summary>
        /// <param name="shapeId">The SHAPE_ID of the record to update.</param>
        /// <param name="fieldName">The name of the column to update.</param>
        /// <param name="value">The new value for the field.</param>
        /// <param name="icType">The current IC Type, used to find the correct table name from the rules.</param>
        public async Task UpdateShapeInfoFieldAsync(string shapeId, string fieldName, object value, string icType)
        {
            const string methodName = "UpdateShapeInfoFieldAsync";

            // 1. A whitelist of columns that are safe to update dynamically.
            var allowedColumns = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
    {
        "SHAPE_STATUS", "CREATED_BY", "CENTROID_X", "CENTROID_Y", "SITE_DIST"
    };

            if (!allowedColumns.Contains(fieldName))
            {
                throw new ArgumentException($"The field '{fieldName}' is not allowed for dynamic updates.");
            }

            // 2. Get the name of the shape info table from the rules engine.
            var shapeInfoTableRule = _rules.ReturnIcGisTypeSettings(icType)?.ShapeInfoTable;
            if (shapeInfoTableRule == null || string.IsNullOrEmpty(shapeInfoTableRule.PostGreFeatureClassName))
            {
                throw new InvalidOperationException($"The 'shape_info_table' is not configured for IC Type '{icType}'.");
            }
            string tableName = shapeInfoTableRule.PostGreFeatureClassName;

            // 3. Build the SQL and parameter list.
            string sql = $"UPDATE {tableName} SET {fieldName} = ? WHERE SHAPE_ID = ?";
            var parameters = new List<object> { value, shapeId };

            try
            {
                await Task.Run(() => PostGreTool.ExecuteRawQuery(sql, parameters, "NOTHING"));
            }
            catch (Exception ex)
            {
                Log.RecordError($"Error updating field '{fieldName}' for Shape ID '{shapeId}'.", ex, methodName);
                throw;
            }
        }

        /// <summary>
        /// Copies an approved shape's geometry and key attributes into the 'proposed' feature class.
        /// </summary>
        /// <param name="shape">The geometry of the shape to save.</param>
        /// <param name="newShapeId">The new, unique SHAPE_ID for the feature.</param>
        /// <param name="icType">The current IC Type, used to find the correct feature class from the rules.</param>
        public async Task CopyShapeToProposedAsync(Geometry shape, string newShapeId, string icType)
        {
            const string methodName = "CopyShapeToProposedAsync";

            // All geodatabase edits must run on the background QueuedTask.
            bool success = await QueuedTask.Run(() =>
            {
                // 1. Get the rule for the 'proposed' feature class from the IC_Rules engine.
                var proposedFcRule = _rules.ReturnIcGisTypeSettings(icType)?.ProposedFeatureClass;
                if (proposedFcRule == null || string.IsNullOrEmpty(proposedFcRule.PostGreFeatureClassName))
                {
                    _log.RecordError("Proposed feature class is not configured in the rules.", null, methodName);
                    return false;
                }

                var workspaceRule = proposedFcRule.WorkSpaceRule;
                Geodatabase geodatabase = null;
                try
                {
                    // 2. Build the connection properties and connect to the geodatabase.
                    var dbConnectionProperties = new DatabaseConnectionProperties(EnterpriseDatabaseType.PostgreSQL)
                    {
                        Instance = $"sde:postgresql:{workspaceRule.Server}",
                        Database = workspaceRule.Database,
                        User = workspaceRule.User,
                        Password = workspaceRule.Password,
                        Version = workspaceRule.Version
                    };
                    geodatabase = new Geodatabase(dbConnectionProperties);

                    // 3. Open the target feature class.
                    using (var proposedFc = geodatabase.OpenDataset<FeatureClass>(proposedFcRule.PostGreFeatureClassName))
                    {
                        // 4. Create an EditOperation to manage the edit.
                        var editOperation = new EditOperation
                        {
                            Name = $"Copy Proposed Shape {newShapeId}",
                            ShowProgressor = false,
                            ShowModalMessageAfterFailure = false
                        };

                        // 5. Prepare the attributes for the new feature.
                        var attributes = new Dictionary<string, object>
                {
                    { "SHAPE_ID", newShapeId },
                    { proposedFc.GetDefinition().GetShapeField(), shape } // Add the geometry
                };

                        // 6. Queue the creation of the new feature.
                        editOperation.Create(proposedFc, attributes);

                        // 7. Execute the operation to commit the new feature to the geodatabase.
                        return editOperation.Execute();
                    }
                }
                catch (Exception ex)
                {
                    _log.RecordError($"Failed to copy shape to the proposed feature class '{proposedFcRule.PostGreFeatureClassName}'.", ex, methodName);
                    return false;
                }
                finally
                {
                    geodatabase?.Dispose();
                }
            });

            if (success)
            {
                Log.RecordMessage($"Shape {newShapeId} successfully copied to proposed feature class.", BisLogMessageType.Note);
            }
            else
            {
                Log.RecordError($"Failed to execute EditOperation for shape {newShapeId}.", null, methodName);
            }
        }

        /// <summary>
        /// Retrieves key business attributes for a single shape from the shape_info table.
        /// </summary>
        /// <param name="shapeId">The SHAPE_ID to look up.</param>
        /// <returns>A nullable tuple containing the PrefId, IcType, and Status, or null if not found.</returns>
        private async Task<(string PrefId, string IcType, string Status)?> GetShapeInfoAsync(string shapeId)
        {
            const string methodName = "GetShapeInfoAsync";
            (string PrefId, string IcType, string Status)? result = null;

            try
            {
                var paramDict = new Dictionary<string, object> { { "SHAPE_ID", shapeId } };

                // Run the database query on a background thread
                var dt = await Task.Run(() =>
                    PostGreTool.ExecuteNamedQuery("returnShapeInfo", paramDict) as DataTable
                );

                if (dt != null && dt.Rows.Count > 0)
                {
                    DataRow row = dt.Rows[0];
                    // Safely extract the values from the DataRow
                    string prefId = row["PREF_ID"]?.ToString() ?? "Unknown";
                    string icType = row["IC_TYPE"]?.ToString() ?? "";
                    string status = row["SHAPE_STATUS"]?.ToString() ?? "";

                    result = (prefId, icType, status);
                }
            }
            catch (Exception ex)
            {
                _log.RecordError($"Error retrieving shape info for Shape ID '{shapeId}'.", ex, methodName);
                // Return null in case of an error
                return null;
            }

            return result;
        }
    }
}
