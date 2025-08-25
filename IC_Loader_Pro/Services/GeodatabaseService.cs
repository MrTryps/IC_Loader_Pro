using ArcGIS.Core.Data;
using BIS_Tools_DataModels_2025;
using System;
using System.Threading.Tasks;
using static IC_Loader_Pro.Module1;

namespace IC_Loader_Pro.Services
{
    /// <summary>
    /// A centralized service for connecting to geodatabases and accessing datasets.
    /// </summary>
    public class GeodatabaseService
    {
        /// <summary>
        /// Connects to an enterprise geodatabase and opens a specific feature class.
        /// </summary>
        public async Task<FeatureClass> GetFeatureClassAsync(FeatureclassRule featureclassRule)
        {
            if (featureclassRule == null || string.IsNullOrEmpty(featureclassRule.PostGreFeatureClassName))
            {
                Log.RecordError("The provided feature class rule is null or does not specify a feature class name.", null, nameof(GetFeatureClassAsync));
                return null;
            }

            Geodatabase geodatabase = null;
            try
            {
                geodatabase = await OpenWorkspaceAsync(featureclassRule.WorkSpaceRule);
                if (geodatabase == null) return null;

                // The 'using' block is not needed here as the GDB connection is managed by the caller of this service
                return geodatabase.OpenDataset<FeatureClass>(featureclassRule.PostGreFeatureClassName);
            }
            catch (Exception ex)
            {
                Log.RecordError($"Failed to open feature class '{featureclassRule.PostGreFeatureClassName}'.", ex, nameof(GetFeatureClassAsync));
                geodatabase?.Dispose();
                return null;
            }
        }


        /// <summary>
        /// Connects to an enterprise geodatabase and opens a specific non-spatial table.
        /// </summary>
        /// <param name="tableRule">The rule containing the workspace and table name information.</param>
        /// <returns>A Table object, or null if an error occurs.</returns>
        public async Task<Table> GetTableAsync(FeatureclassRule tableRule)
        {
            if (tableRule == null || string.IsNullOrEmpty(tableRule.PostGreFeatureClassName))
            {
                Log.RecordError("The provided table rule is null or does not specify a table name.", null, nameof(GetTableAsync));
                return null;
            }

            Geodatabase geodatabase = null;
            try
            {
                geodatabase = await OpenWorkspaceAsync(tableRule.WorkSpaceRule);
                if (geodatabase == null) return null;

                return geodatabase.OpenDataset<Table>(tableRule.PostGreFeatureClassName);
            }
            catch (Exception ex)
            {
                Log.RecordError($"Failed to open table '{tableRule.PostGreFeatureClassName}'.", ex, nameof(GetTableAsync));
                geodatabase?.Dispose();
                return null;
            }
        }

        /// <summary>
        /// Private helper to connect to a workspace based on a WorkSpaceRule.
        /// </summary>
        private Task<Geodatabase> OpenWorkspaceAsync(WorkSpaceRule workspaceRule)
        {
            if (workspaceRule == null)
            {
                Log.RecordError("Workspace rule cannot be null.", null, nameof(OpenWorkspaceAsync));
                return Task.FromResult<Geodatabase>(null);
            }

            try
            {
                string instance = workspaceRule.Server; 
                if (workspaceRule.Port >0)                
                {
                    instance += $", {workspaceRule.Port}"; // Append port if specified
                }
                var dbConnectionProperties = new DatabaseConnectionProperties(EnterpriseDatabaseType.PostgreSQL)
                {
                    AuthenticationMode = AuthenticationMode.DBMS,
                    Instance = instance,// $"sde:postgresql:{workspaceRule.Server}",
                    Database = workspaceRule.Database,
                    User = workspaceRule.User,
                    Password = workspaceRule.Password,
                    Version = workspaceRule.Version
                };

                return Task.FromResult(new Geodatabase(dbConnectionProperties));
            }
            catch (Exception ex)
            {
                Log.RecordError($"Failed to create geodatabase connection properties for workspace '{workspaceRule.WorkspaceName}'.", ex, nameof(OpenWorkspaceAsync));
                return Task.FromResult<Geodatabase>(null);
            }
        }
    }
}