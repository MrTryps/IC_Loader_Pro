using ArcGIS.Core.Data;
using ArcGIS.Core.Geometry;
using ArcGIS.Desktop.Framework.Threading.Tasks;
using BIS_Tools_DataModels_2025;
using IC_Loader_Pro.Models;
using IC_Rules_2025;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using static BIS_Log;
using Path = System.IO.Path;
using QueryFilter = ArcGIS.Core.Data.QueryFilter;

namespace IC_Loader_Pro.Services
{
    /// <summary>
    /// A service dedicated to reading, validating, and repairing features (shapes)
    /// from identified GIS filesets.
    /// </summary>
    public class FeatureProcessingService
    {
        private readonly BIS_Log _log;
        private readonly IC_Rules _rules;
        private readonly IcNamedTests _namedTests;

        public FeatureProcessingService(IC_Rules rules, IcNamedTests namedTests, BIS_Log log)
        {
            _rules = rules;
            _namedTests = namedTests;
            _log = log;
        }

        /// <summary>
        /// (SHELL METHOD) The main entry point for orchestrating the analysis of all features
        /// from a list of valid filesets.
        /// </summary>
        /// <param name="validFilesets">The list of valid filesets identified from the email attachments.</param>
        /// <returns>A list of ShapeItem objects ready for UI display and user review.</returns>
        public async Task<List<ShapeItem>> AnalyzeFeaturesFromFilesetsAsync(List<fileset> identifiedFileSets, string icType, MapPoint siteLocation)
        {
            _log.RecordMessage($"Starting feature analysis for {identifiedFileSets.Count} submitted fileset(s)...", BIS_Log.BisLogMessageType.Note);
            var allAnalyzedShapes = new List<ShapeItem>();

            if (identifiedFileSets == null || !identifiedFileSets.Any())
            {
                return allAnalyzedShapes; // Return an empty list if there's nothing to process
            }

            // Loop through each valid fileset that was identified.
            foreach (var fileset in identifiedFileSets.Where(fs => fs.validSet))
            {
                // 1. Read all the raw features from the current fileset (e.g., a shapefile).
                List<ShapeItem> shapesFromFile = await ReadFeaturesFromFileAsync(fileset, icType);

                if (!shapesFromFile.Any())
                {
                    // If the file was empty or unreadable, continue to the next fileset.
                    continue;
                }

                // 2. Loop through each feature found in the file and run our validation checks.
                foreach (var shapeItem in shapesFromFile)
                {
                    // This method will perform all the checks (projection, self-intersection, area, etc.)
                    // and update the shapeItem's IsValid and Status properties.
                    ValidateShape(shapeItem, icType, siteLocation);

                    // Add the fully analyzed shape to our master list.
                    allAnalyzedShapes.Add(shapeItem);
                }
            }

            _log.RecordMessage($"Feature analysis complete. A total of {allAnalyzedShapes.Count} shapes were extracted and analyzed.", BIS_Log.BisLogMessageType.Note);
            return allAnalyzedShapes;
        }

        private void ValidateShape(ShapeItem shapeToValidate, string icType, MapPoint siteLocation)
        {
            if (shapeToValidate?.Geometry == null)
            {
                shapeToValidate.IsValid = false;
                shapeToValidate.Status = "Missing Geometry";
                return;
            }

            // Get the geometry rules for the current IC Type
            var geometryRules = _rules.ReturnIcGisTypeSettings(icType)?.GeometryRules;
            if (geometryRules == null)
            {
                shapeToValidate.IsValid = false;
                shapeToValidate.Status = "Missing Geometry Rules";
                return;
            }

            var geometry = shapeToValidate.Geometry;

            // 1: Check and Reproject Spatial Reference ---
            // We will use the WKID for NAD 1983 State Plane New Jersey FIPS 2900 (US Feet).
            // The modern WKID is 102711 (the older one was 2260).
            var njspfSr = SpatialReferenceBuilder.CreateSpatialReference(2260);
            if (geometry.SpatialReference == null || !geometry.SpatialReference.IsEqual(njspfSr))
            {
                try
                {
                    // If not, reproject it.
                    var projectedGeometry = GeometryEngine.Instance.Project(geometry, njspfSr);
                    if (projectedGeometry != null)
                    {
                        shapeToValidate.Geometry = projectedGeometry as Polygon; // Update the geometry
                        _log.RecordMessage($"Shape {shapeToValidate.ShapeReferenceId} was reprojected to NJ State Plane.", BisLogMessageType.Note);
                    }
                    else
                    {
                        throw new Exception("Projection returned a null geometry.");
                    }
                }
                catch (Exception ex)
                {
                    _log.RecordError($"Failed to reproject geometry for shape {shapeToValidate.ShapeReferenceId}.", ex, "ValidateShape");
                    shapeToValidate.IsValid = false;
                    shapeToValidate.Status = "Reprojection Failed";
                    return; // Stop validation if reprojection fails.
                }
            }

            // 2. Check if the shape is a polygon
            if (geometry.GeometryType != GeometryType.Polygon)
            {
                shapeToValidate.IsValid = false;
                shapeToValidate.Status = $"Invalid Type: {geometry.GeometryType}";
                return;
            }

            // 3. Check if the geometry is empty
            if (geometry.IsEmpty)
            {
                shapeToValidate.IsValid = false;
                shapeToValidate.Status = "Empty Geometry";
                return;
            }

            // 4. Check and correct the spatial reference (projection)
            var requiredSr = SpatialReferenceBuilder.CreateSpatialReference(geometryRules.ProjectionId);
            if (!geometry.SpatialReference.IsEqual(requiredSr))
            {
                try
                {
                    geometry = (Polygon)GeometryEngine.Instance.Project(geometry, requiredSr);
                    shapeToValidate.Geometry = geometry as Polygon; // Update the geometry
                }
                catch (Exception ex)
                {
                    _log.RecordError("Failed to reproject geometry.", ex, "ValidateShape");
                    shapeToValidate.IsValid = false;
                    shapeToValidate.Status = "Reprojection Failed";
                    return;
                }
            }

            // 5. Check for self-intersection and simplify if necessary (the modern way)
            // The GeometryEngine's SimplifyAsFeature method can fix many common geometry errors.
            if (!GeometryEngine.Instance.IsSimpleAsFeature(geometry))
            {
                geometry = (Polygon)GeometryEngine.Instance.SimplifyAsFeature(geometry, true); // true = allow endpoint changes
                shapeToValidate.Geometry = geometry as Polygon; // Update the geometry
                shapeToValidate.Status = "Repaired (Simplified)";
            }

            // 6. Check the area against the minimum threshold
            double area = (geometry as Polygon).Area;
            shapeToValidate.Area = area;
            if (Math.Abs(area) < geometryRules.Min_Area)
            {
                shapeToValidate.IsValid = false;
                shapeToValidate.Status = "Area Below Minimum";
                return;
            }

            // 7.  Check if the shape is within the allowed extent ---
            var extent = geometry.Extent;
            if (extent.XMin < geometryRules.X_Min ||
                extent.XMax > geometryRules.X_Max ||
                extent.YMin < geometryRules.Y_Min ||
                extent.YMax > geometryRules.Y_Max)
            {
                shapeToValidate.IsValid = false;
                shapeToValidate.Status = "Outside Allowable Extent";
                return;
            }

            // 8. Check the distance from the site, if a site location was provided.
            if (siteLocation != null && shapeToValidate.Geometry != null)
            {
                //_log.RecordMessage("--- Preparing for Distance Calculation ---", BisLogMessageType.Note);
                //_log.RecordMessage($"Polygon SR WKID: {geometry.SpatialReference?.Wkid ?? -1}", BisLogMessageType.Note);
                //_log.RecordMessage($"Site Location SR WKID: {siteLocation.SpatialReference?.Wkid ?? -1}", BisLogMessageType.Note);
                //_log.RecordMessage($"Required SR WKID: {njspfSr.Wkid}", BisLogMessageType.Note);
                // Use the GeometryEngine to calculate the geodetic distance.
                double distance = GeometryEngine.Instance.Distance(shapeToValidate.Geometry, siteLocation);
                shapeToValidate.DistanceFromSite = distance; // Store the distance

                if (geometryRules != null && distance > geometryRules.SiteDistance)
                {
                    shapeToValidate.IsValid = false;
                    shapeToValidate.Status = "Exceeds Max Distance from Site";
                }
            }

            // If all checks pass, the shape is considered valid
            shapeToValidate.IsValid = true;
            if (shapeToValidate.Status == "Pending Validation") // Only update if not already repaired
            {
                shapeToValidate.Status = "Valid";
            }
        }

        /// <summary>
        /// Reads all features from a single GIS file (e.g., a shapefile) and converts them
        /// into a list of ShapeItem objects.
        /// </summary>
        /// <param name="fileset">The fileset to read from.</param>
        /// <param name="icType">The IC Type being processed, to get the correct fields to mine.</param>
        /// <returns>A list of ShapeItem objects, one for each feature in the file.</returns>
        private async Task<List<ShapeItem>> ReadFeaturesFromFileAsync(fileset fileset, string icType)
        {
            var shapesInFile = new List<ShapeItem>();
            _log.RecordMessage($"Reading features from file: {fileset.fileName}", BIS_Log.BisLogMessageType.Note);

            // This entire block of GIS code MUST run on the ArcGIS Pro background thread (MCT).
            await QueuedTask.Run(() =>
            {
                // Construct the full path to the .shp file
                string shpPath = Path.Combine(fileset.path, fileset.fileName + ".shp");

                if (!File.Exists(shpPath))
                {
                    _log.RecordError($"Shapefile not found at expected path: {shpPath}", null, "ReadFeaturesFromFileAsync");
                    return; // Exit if the shapefile doesn't exist
                }

                try
                {
                    // The ArcGIS Pro SDK connects to a folder containing shapefiles as if it were a geodatabase.
                    var connectionPath = new FileSystemConnectionPath(new Uri(fileset.path), FileSystemDatastoreType.Shapefile);
                    using (var datastore = new FileSystemDatastore(connectionPath))
                    using (var featureClass = datastore.OpenDataset<FeatureClass>(fileset.fileName))
                    {
                        if (featureClass == null) return;

                        // Get the list of attribute fields we need to extract from the rules engine.
                        var fieldsToMine = _rules.ReturnIcGisTypeSettings(icType)
                                                 .FeatureFields
                                                 .Where(f => f.DisplayInPreview)
                                                 .Select(f => f.Fieldname)
                                                 .ToList();

                        // Use a QueryFilter to specify which fields we want to retrieve.
                        var queryFilter = new QueryFilter { SubFields = "*" }; // For now, get all fields

                        // Loop through each feature (row) in the feature class.
                        using (var cursor = featureClass.Search(queryFilter, false))
                        {
                            while (cursor.MoveNext())
                            {
                                using (var feature = cursor.Current as Feature)
                                {
                                    // We are only interested in Polygon features.
                                    if (feature?.GetShape() is Polygon polygon)
                                    {
                                        var shapeItem = new ShapeItem
                                        {
                                            Geometry = polygon,
                                            SourceFile = fileset.fileName,
                                            ShapeReferenceId = (int)feature.GetObjectID(),
                                            ShapeType = feature.GetShape().GeometryType.ToString(),
                                            IsValid = true, // Assume valid for now; the ValidateShape method will check this later.
                                            Status = "Pending Validation"
                                        };

                                        // Extract the attribute values for the "fields to mine".
                                        foreach (string fieldName in fieldsToMine)
                                        {
                                            int fieldIndex = feature.FindField(fieldName);
                                            if (fieldIndex != -1)
                                            {
                                                shapeItem.Attributes[fieldName] = feature[fieldIndex];
                                            }
                                        }
                                        shapesInFile.Add(shapeItem);
                                    }
                                }
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    _log.RecordError($"An unexpected error occurred while reading features from '{shpPath}'.", ex, "ReadFeaturesFromFileAsync");
                    // In a real-world scenario, you might create a "failed" ShapeItem here to notify the user.
                }
            });

            return shapesInFile;
        }

        /// <summary>
        /// (SHELL METHOD) Performs a series of validation checks on a single shape's geometry.
        /// </summary>
        /// <param name="shapeToValidate">The ShapeItem to validate.</param>
        private void ValidateShape(ShapeItem shapeToValidate)
        {
            // --- FUTURE LOGIC ---
            // This method will contain all the business logic checks from your reference code:
            // - Is it a polygon?
            // - Is it empty?
            // - Is it in the correct projection? (and reproject if not)
            // - Is it self-intersecting?
            // - Is its area above the minimum threshold?
            // It will update the shapeToValidate.IsValid and shapeToValidate.Status properties.
        }
    }
}