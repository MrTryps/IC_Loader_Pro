using ArcGIS.Core.Data;
using ArcGIS.Core.Geometry;
using ArcGIS.Desktop.Core;
using ArcGIS.Desktop.Editing;
using ArcGIS.Desktop.Framework.Threading.Tasks;
using IC_Loader_Pro.Models;
using IC_Rules_2025;
using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Runtime.CompilerServices;
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
        //public async Task<bool> RecordShapeInfoAsync(string newShapeId, string submissionId, string deliverableId, string prefId, string icType)
        //{
        //    const string methodName = "RecordShapeInfoAsync";

        //    // 1. Get the rule for the shape info table from the rules engine.
        //    var shapeInfoTableRule = _rules.ReturnIcGisTypeSettings(icType)?.ShapeInfoTable;
        //    if (shapeInfoTableRule == null)
        //    {
        //        _log.RecordError($"The 'shape_info_table' is not configured for IC Type '{icType}'.", null, methodName);
        //        return false;
        //    }

        //    // 2. All geodatabase edits must be run on the QueuedTask.
        //    bool success = await QueuedTask.Run(async () =>
        //    {
        //        try
        //        {
        //            var gdbService = new GeodatabaseService();
        //            // 3. Use our new service to open the table.
        //            using (ArcGIS.Core.Data.Table shapeInfoTable = await gdbService.GetTableAsync(shapeInfoTableRule))
        //            {
        //                if (shapeInfoTable == null) return false;

        //                // 4. Create and execute an EditOperation to create the new row.
        //                var editOperation = new EditOperation
        //                {
        //                    Name = $"Create record for {newShapeId}",
        //                    ShowProgressor = false,
        //                    ShowModalMessageAfterFailure = false
        //                };

        //                // 5. Prepare the attributes for the new row.
        //                var attributes = new Dictionary<string, object>
        //        {
        //            { "shape_id", newShapeId },
        //            { "submission_id", submissionId },
        //            { "deliverable_id", deliverableId },
        //            { "pref_id", prefId },
        //            { "ic_type", icType }
        //        };

        //                // 6. Queue the creation of the new row.
        //                editOperation.Create(shapeInfoTable, attributes);
        //                bool wasExecuted = editOperation.Execute();

        //                // --- THIS IS THE NEW ERROR-LOGGING LOGIC ---
        //                if (!wasExecuted)
        //                {
        //                    // If the operation fails, log the specific error message from the EditOperation.
        //                    Log.RecordError($"EditOperation failed to create shape info record. Reason: {editOperation.ErrorMessage}", null, methodName);
        //                }
        //                return wasExecuted;
        //            }
        //        }
        //        catch (Exception ex)
        //        {
        //            _log.RecordError($"Error creating new shape info record in '{shapeInfoTableRule.PostGreFeatureClassName}' for Shape ID '{newShapeId}'.", ex, methodName);
        //            return false;
        //        }
        //    });

        //    if (success)
        //    {
        //        Log.RecordMessage($"Successfully created record in shape_info table for Shape ID: {newShapeId}", BisLogMessageType.Note);
        //    }
        //    else
        //    {
        //        // The specific error will have already been logged inside the QueuedTask.
        //    }
        //    return success;
        //}

        // In IC_Loader_Pro/Services/ShapeProcessingService.cs

        //public async Task<bool> RecordShapeInfoAsync(string newShapeId, string submissionId, string deliverableId, string prefId, string icType)
        //{
        //    const string methodName = "RecordShapeInfoAsync";

        //    var shapeInfoTableRule = IcRules.ReturnIcGisTypeSettings(icType)?.ShapeInfoTable;
        //    if (shapeInfoTableRule == null || string.IsNullOrEmpty(shapeInfoTableRule.PostGreFeatureClassName))
        //    {
        //        Log.RecordError($"The 'shape_info_table' is not configured for IC Type '{icType}'.", null, methodName);
        //        return false;
        //    }
        //    string tableName = shapeInfoTableRule.PostGreFeatureClassName;

        //    try
        //    {
        //        // --- THIS IS THE CORRECTED LINE ---
        //        // Pass the service's own logger (_log) into the factory method.
        //        var shapeInfoDbTool = BIS_DB_PostGre.CreateFromRule(shapeInfoTableRule.WorkSpaceRule, _log);
        //        if (shapeInfoDbTool == null)
        //        {
        //            Log.RecordError("Failed to create a database connection for the shape_info table.", null, methodName);
        //            return false;
        //        }

        //        var attributes = new Dictionary<string, object> { /*...*/ };
        //        var fieldList = string.Join(", ", attributes.Keys);
        //        var valuePlaceholders = string.Join(", ", attributes.Keys.Select(k => "?"));
        //        var sql = $"INSERT INTO {tableName} ({fieldList}) VALUES ({valuePlaceholders})";
        //        var parameters = attributes.Values.ToList();

        //        await Task.Run(() => shapeInfoDbTool.ExecuteRawQuery(sql, parameters, "NOTHING"));

        //        Log.RecordMessage($"Successfully created record in {tableName} for Shape ID: {newShapeId}", BisLogMessageType.Note);
        //        return true;
        //    }
        //    catch (Exception ex)
        //    {
        //        Log.RecordError($"Error creating new shape info record in '{tableName}' for Shape ID '{newShapeId}'.", ex, methodName);
        //        return false;
        //    }
        //}

        //public async Task<bool> RecordShapeInfoAsync(string newShapeId, string submissionId, string deliverableId, string prefId, string icType)
        //{
        //    const string methodName = "RecordShapeInfoAsync";

        //    var shapeInfoTableRule = IcRules.ReturnIcGisTypeSettings(icType)?.ShapeInfoTable;
        //    if (shapeInfoTableRule == null || string.IsNullOrEmpty(shapeInfoTableRule.PostGreFeatureClassName))
        //    {
        //        Log.RecordError($"The 'shape_info_table' is not configured for IC Type '{icType}'.", null, methodName);
        //        return false;
        //    }

        //    bool success = await QueuedTask.Run(async () =>
        //    {
        //        try
        //        {
        //            var gdbService = new GeodatabaseService();
        //            using (ArcGIS.Core.Data.Table shapeInfoTable = await gdbService.GetTableAsync(shapeInfoTableRule))
        //            {
        //                if (shapeInfoTable == null) return false;

        //                var editOperation = new EditOperation
        //                {
        //                    Name = $"Create record for {newShapeId}",
        //                    ShowProgressor = false,
        //                    ShowModalMessageAfterFailure = false
        //                };

        //                // The insert logic is queued as a Callback action within the EditOperation.
        //                // This is the most robust pattern.
        //                editOperation.Callback(context =>
        //                {
        //                    // Create a RowBuffer to hold the new row's attributes.
        //                    using (var rowBuffer = shapeInfoTable.CreateRowBuffer())
        //                    {
        //                        rowBuffer["SHAPE_ID"] = newShapeId;
        //                        rowBuffer["submission_id"] = submissionId;
        //                        rowBuffer["deliverable_id"] = deliverableId;
        //                        rowBuffer["pref_id"] = prefId;
        //                        rowBuffer["ic_type"] = icType;

        //                        // Create the new row using an InsertCursor.
        //                        using (var cursor = shapeInfoTable.CreateInsertCursor())
        //                        {
        //                            cursor.Insert(rowBuffer);
        //                        }
        //                    }
        //                }, shapeInfoTable);

        //                bool wasExecuted = editOperation.Execute();
        //                if (!wasExecuted)
        //                {
        //                    Log.RecordError($"EditOperation failed to create shape info record. Reason: {editOperation.ErrorMessage}", null, methodName);
        //                }
        //                return wasExecuted;
        //            }
        //        }
        //        catch (Exception ex)
        //        {
        //            Log.RecordError($"Error creating new shape info record for Shape ID '{newShapeId}'.", ex, methodName);
        //            return false;
        //        }
        //    });

        //    if (success)
        //    {
        //        Log.RecordMessage($"Successfully created record in shape_info table for Shape ID: {newShapeId}", BisLogMessageType.Note);
        //    }

        //    return success;
        //}

        public async Task<bool> RecordShapeInfoAsync(string newShapeId, string submissionId, string deliverableId, string prefId, string icType)
        {
            const string methodName = "RecordShapeInfoAsync";

            var shapeInfoTableRule = IcRules.ReturnIcGisTypeSettings(icType)?.ShapeInfoTable;
            if (shapeInfoTableRule == null || string.IsNullOrEmpty(shapeInfoTableRule.PostGreFeatureClassName))
            {
                Log.RecordError($"The 'shape_info_table' is not configured for IC Type '{icType}'.", null, methodName);
                return false;
            }
            string tableName = shapeInfoTableRule.PostGreFeatureClassName;

            try
            {
                // 1. Create a dedicated database tool that connects to the correct cluster for the shape_info table.
                var shapeInfoDbTool = BIS_DB_PostGre.CreateFromRule(shapeInfoTableRule.WorkSpaceRule, _log);
                if (shapeInfoDbTool == null)
                {
                    Log.RecordError("Failed to create a database connection for the shape_info table.", null, methodName);
                    return false;
                }

                // 2. Build the dynamic INSERT statement and parameter list.
                var attributes = new Dictionary<string, object>
        {
            { "SHAPE_ID", newShapeId },
            { "submission_id", submissionId },
            { "deliverable_id", deliverableId },
            { "pref_id", prefId },
            { "ic_type", icType }
        };

                var fieldList = string.Join(", ", attributes.Keys);
                var valuePlaceholders = string.Join(", ", attributes.Keys.Select(k => "?"));
                var sql = $"INSERT INTO {tableName} ({fieldList}) VALUES ({valuePlaceholders})";
                var parameters = attributes.Values.ToList();

                // 3. Execute the query using the PostGreTool.
                await Task.Run(() => shapeInfoDbTool.ExecuteRawQuery(sql, parameters, "NOTHING"));

                Log.RecordMessage($"Successfully created record in {tableName} for Shape ID: {newShapeId}", BisLogMessageType.Note);
                return true;
            }
            catch (Exception ex)
            {
                Log.RecordError($"Error creating new shape info record in '{tableName}' for Shape ID '{newShapeId}'.", ex, methodName);
                return false;
            }
        }



        /// <summary>
        /// Checks for duplicate polygons in the proposed feature class by comparing both
        /// geometry (with a tolerance) and key business attributes.
        /// </summary>
        public async Task<bool> IsDuplicateInProposedAsync(Geometry shape, string currentPrefId, string icType)
        {
            bool isDuplicate = false;

            var proposedFcRule = _rules.ReturnIcGisTypeSettings(icType)?.ProposedFeatureClass;
            if (proposedFcRule == null)
            {
                _log.RecordError("Proposed feature class is not configured in the rules.", null, nameof(IsDuplicateInProposedAsync));
                return false;
            }

            await QueuedTask.Run(async () =>
            {
                try 
                {
                    var gdbService = new GeodatabaseService();
                    using (var proposedFc = await gdbService.GetFeatureClassAsync(proposedFcRule))
                    {
                        if (proposedFc == null) return;

                        var spatialFilter = new SpatialQueryFilter
                        {
                            FilterGeometry = shape.Extent,
                            SpatialRelationship = SpatialRelationship.Intersects
                        };

                        using (var cursor = proposedFc.Search(spatialFilter, false))
                        {
                            while (cursor.MoveNext())
                            {
                                using (var currentFeature = cursor.Current as Feature)
                                {
                                    if (currentFeature == null) continue;
                                    var existingShape = currentFeature.GetShape();
                                    if (existingShape == null) continue;
                                    var shapeToCompare = GeometryEngine.Instance.Project(shape, existingShape.SpatialReference);
                                    if (GeometryEngine.Instance.Equals(shape, currentFeature.GetShape()))
                                    {
                                        string existingShapeId = currentFeature["SHAPE_ID"]?.ToString();
                                        if (string.IsNullOrEmpty(existingShapeId)) continue;

                                        var existingShapeInfo = await GetShapeInfoAsync(existingShapeId, icType);

                                        if (existingShapeInfo.HasValue &&
                                            existingShapeInfo.Value.PrefId.Equals(currentPrefId, StringComparison.OrdinalIgnoreCase) &&
                                            existingShapeInfo.Value.IcType.Equals(icType, StringComparison.OrdinalIgnoreCase) &&
                                            (existingShapeInfo.Value.Status == "To Be Reviewed" || existingShapeInfo.Value.Status == "Shape Approved"))
                                        {
                                            _log.RecordMessage($"Found a valid duplicate of existing shape ID: {existingShapeId}", BisLogMessageType.Note);
                                            isDuplicate = true;
                                            return;
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    _log.RecordError($"A critical error occurred while checking for duplicate shapes in {proposedFcRule.PostGreFeatureClassName}.", ex, nameof(IsDuplicateInProposedAsync));
                    // In case of an error, we assume it's not a duplicate to be safe.
                    isDuplicate = false;
                }
            });

            return isDuplicate;
        }


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
                Log.RecordError($"The field '{fieldName}' is not allowed for dynamic updates.", new ArgumentException(fieldName), methodName);
                return;
            }

            // 2. Get the rule for the shape info table from the rules engine.
            var shapeInfoTableRule = IcRules.ReturnIcGisTypeSettings(icType)?.ShapeInfoTable;
            if (shapeInfoTableRule == null || string.IsNullOrEmpty(shapeInfoTableRule.PostGreFeatureClassName))
            {
                Log.RecordError($"The 'shape_info_table' is not configured for IC Type '{icType}'.", null, methodName);
                return;
            }
            string tableName = shapeInfoTableRule.PostGreFeatureClassName;

            try
            {
                // 3. Create a dedicated database tool that connects to the correct cluster.
                var shapeInfoDbTool = BIS_DB_PostGre.CreateFromRule(shapeInfoTableRule.WorkSpaceRule, _log);
                if (shapeInfoDbTool == null)
                {
                    Log.RecordError("Failed to create a database connection for the shape_info table.", null, methodName);
                    return;
                }

                // 4. Build the SQL UPDATE statement and parameter list.
                string sql = $"UPDATE {tableName} SET {fieldName} = ? WHERE SHAPE_ID = ?";
                var parameters = new List<object> { value, shapeId };

                // 5. Execute the query using the PostGreTool.
                await Task.Run(() => shapeInfoDbTool.ExecuteRawQuery(sql, parameters, "NOTHING"));
            }
            catch (Exception ex)
            {
                Log.RecordError($"Error updating field '{fieldName}' for Shape ID '{shapeId}'.", ex, methodName);
                // Decide if you want to re-throw the exception to stop the save process
                // throw; 
            }
        }

        // In IC_Loader_Pro/Services/ShapeProcessingService.cs

        public async Task CopyShapeToProposedAsync(Geometry shape, string newShapeId, string icType)
        {
            const string methodName = "CopyShapeToProposedAsync";

            var proposedFcRule = IcRules.ReturnIcGisTypeSettings(icType)?.ProposedFeatureClass;
            if (proposedFcRule == null)
            {
                Log.RecordError("Proposed feature class is not configured in the rules.", null, methodName);
                return;
            }

            bool success = await QueuedTask.Run(async () =>
            {
                var gdbService = new GeodatabaseService();
                using (var proposedFc = await gdbService.GetFeatureClassAsync(proposedFcRule))
                {
                    if (proposedFc == null) return false;

                    var editOperation = new EditOperation
                    {
                        Name = $"Copy Proposed Shape {newShapeId}",
                        ShowProgressor = false,
                        ShowModalMessageAfterFailure = true
                    };

                    // --- START OF THE FIX ---
                    // This is the standard, compatible pattern for creating a feature.
                    var attributes = new Dictionary<string, object>
            {
                { "SHAPE_ID", newShapeId },
                { proposedFc.GetDefinition().GetShapeField(), shape }
            };

                    // Queue the create operation.
                    editOperation.Create(proposedFc, attributes);

                    // Execute the operation. This now calls the correct parameterless overload.
                    bool wasExecuted = editOperation.Execute();
                    // --- END OF THE FIX ---

                    if (!wasExecuted)
                    {
                        Log.RecordError($"EditOperation failed to copy shape '{newShapeId}'. Reason: {editOperation.ErrorMessage}", null, methodName);
                    }
                    return wasExecuted;
                }
            });

            if (success)
            {
                Log.RecordMessage($"Shape {newShapeId} successfully copied to proposed feature class.", BisLogMessageType.Note);
                await Project.Current.SaveEditsAsync();
            }
            else
            {
                ArcGIS.Desktop.Framework.Dialogs.MessageBox.Show(
                    $"Failed to save the new shape '{newShapeId}' to the geodatabase. Please check the logs for details.",
                    "Save Error");
            }
        }



        /// <summary>
        /// Copies an approved shape's geometry and key attributes into the 'proposed' feature class.
        /// </summary>
        /// <param name="shape">The geometry of the shape to save.</param>
        /// <param name="newShapeId">The new, unique SHAPE_ID for the feature.</param>
        /// <param name="icType">The current IC Type, used to find the correct feature class from the rules.</param>
        public async Task CopyShapeToProposedAsync_bak(Geometry shape, string newShapeId, string icType)
        {
            const string methodName = "CopyShapeToProposedAsync";

            var proposedFcRule = IcRules.ReturnIcGisTypeSettings(icType)?.ProposedFeatureClass;
            if (proposedFcRule == null)
            {
                Log.RecordError("Proposed feature class is not configured in the rules.", null, methodName);
                return;
            }

            bool success = await QueuedTask.Run(async () =>
            {
                var editOperation = new EditOperation
                {
                    Name = $"Copy Proposed Shape {newShapeId}",
                    ShowProgressor = false,
                    ShowModalMessageAfterFailure = false
                };

                try
                {
                    var gdbService = new GeodatabaseService();
                    using (var proposedFc = await gdbService.GetFeatureClassAsync(proposedFcRule))
                    {
                        if (proposedFc == null) return false;

                        var attributes = new Dictionary<string, object>
                {
                    { "SHAPE_ID", newShapeId },
                    { proposedFc.GetDefinition().GetShapeField(), shape }
                };

                        editOperation.Create(proposedFc, attributes);

                        bool wasExecuted = editOperation.Execute();

                        if (!wasExecuted)
                        {
                            // This is the core error logging that will give us the reason for the failure.
                            Log.RecordError($"EditOperation failed to copy shape '{newShapeId}'. Reason: {editOperation.ErrorMessage}", null, methodName);
                        }
                        return wasExecuted;
                    }
                }
                catch (Exception ex)
                {
                    Log.RecordError($"An exception occurred while trying to copy shape '{newShapeId}'.", ex, methodName);
                    // It's always safe to call Abort in a catch block to ensure the operation is cancelled.
                    editOperation.Abort();
                    return false;
                }
            });

            if (success)
            {
                Log.RecordMessage($"Shape {newShapeId} successfully copied to proposed feature class.", BisLogMessageType.Note);
            }
            //StatusMessage = "Saving features to geodatabase...";
            bool editsSaved = await Project.Current.SaveEditsAsync();
            if (!editsSaved)
            {
                Log.RecordError("Failed to save edits to the geodatabase.", null, methodName);
                ArcGIS.Desktop.Framework.Dialogs.MessageBox.Show(
                    "Failed to save new features to the geodatabase. The submission will not be finalized.",
                    "Save Error");
               // IsEmailActionEnabled = true; // Re-enable buttons
                return;
            }
        }

        /// <summary>
        /// Retrieves key business attributes for a single shape from the shape_info table using the Geodatabase API.
        /// </summary>
        private async Task<(string PrefId, string IcType, string Status)?> GetShapeInfoAsync(string shapeId, string icType)
        {
            const string methodName = "GetShapeInfoAsync";
            (string PrefId, string IcType, string Status)? result = null;

            // 1. Get the rule for the shape_info table.
            var shapeInfoTableRule = _rules.ReturnIcGisTypeSettings(icType)?.ShapeInfoTable;
            if (shapeInfoTableRule == null)
            {
                _log.RecordError($"The 'shape_info_table' is not configured for IC Type '{icType}'.", null, methodName);
                return null;
            }

            // This operation must run on the background thread.
            await QueuedTask.Run(async () =>
            {
                try
                {
                    // 2. Use our service to open the table.
                    var gdbService = new GeodatabaseService();
                    using (var shapeInfoTable = await gdbService.GetTableAsync(shapeInfoTableRule))
                    {
                        if (shapeInfoTable == null) return;

                        // 3. Find the specific row using a QueryFilter.
                        var queryFilter = new QueryFilter { WhereClause = $"SHAPE_ID = '{shapeId}'" };
                        using (var cursor = shapeInfoTable.Search(queryFilter, false))
                        {
                            if (cursor.MoveNext())
                            {
                                using (var row = cursor.Current)
                                {
                                    // 4. Extract the values and build the result tuple.
                                    string prefId = row["PREF_ID"]?.ToString() ?? "Unknown";
                                    string foundIcType = row["IC_TYPE"]?.ToString() ?? "";
                                    string status = row["SHAPE_STATUS"]?.ToString() ?? "";
                                    result = (prefId, foundIcType, status);
                                }
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    _log.RecordError($"Error retrieving shape info for Shape ID '{shapeId}'.", ex, methodName);
                    result = null;
                }
            });

            return result;
        }
    }
}
