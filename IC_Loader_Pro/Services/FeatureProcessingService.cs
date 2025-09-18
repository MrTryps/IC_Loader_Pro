using ArcGIS.Core.Data;
using ArcGIS.Core.Geometry;
using ArcGIS.Desktop.Editing;
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
using System.Text.RegularExpressions;
using static IC_Loader_Pro.Module1;

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
        public async Task<List<ShapeItem>> AnalyzeFeaturesFromFilesetsAsync(List<fileset> identifiedFileSets, string icType, MapPoint siteLocation, IcTestResult rootTestResult)
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
                List<ShapeItem> shapesFromFile = await ReadFeaturesFromFileAsync(fileset, icType, rootTestResult);

                if (!shapesFromFile.Any())
                {
                    // If the file was empty or unreadable, continue to the next fileset.
                    IcTestResult noShapesFound = _namedTests.returnNewTestResult("GIS_No_Shapes_Found",fileset.fileName, IcTestResult.TestType.Submission);
                    noShapesFound.Passed = false;
                    rootTestResult.AddSubordinateTestResult(noShapesFound);
                    continue;
                }

                // 2. Loop through each feature found in the file and run our validation checks.
                foreach (var shapeItem in shapesFromFile)
                {
                    // This method will perform all the checks (projection, self-intersection, area, etc.)
                    // and update the shapeItem's IsValid and Status properties.
                    ValidateShape(shapeItem, icType, siteLocation, rootTestResult);

                    // Add the fully analyzed shape to our master list.
                    allAnalyzedShapes.Add(shapeItem);
                }
            }

            _log.RecordMessage($"Feature analysis complete. A total of {allAnalyzedShapes.Count} shapes were extracted and analyzed.", BIS_Log.BisLogMessageType.Note);
            return allAnalyzedShapes;
        }

        private void ValidateShape(ShapeItem shapeToValidate, string icType, MapPoint siteLocation, IcTestResult parentTestResult)
        {
            Action<string> recordShapeCheckFailure = (failureReason) =>
            {
                var shapeCheckTest = _namedTests.returnNewTestResult("GIS_ShapeCheck", shapeToValidate.SourceFile, IcTestResult.TestType.Shape);
                shapeCheckTest.Passed = false;
                shapeCheckTest.AddComment($"Shape with original ID {shapeToValidate.ShapeReferenceId} failed validation: {failureReason}");
                parentTestResult.AddSubordinateTestResult(shapeCheckTest);
                parentTestResult.Passed = false;
            };

            if (shapeToValidate?.Geometry == null)
            {
                shapeToValidate.IsValid = false;
                shapeToValidate.Status = "Missing Geometry";
                recordShapeCheckFailure("Shape feature was found, but its geometry is null.");
                return;
            }

            bool isShapeCurrentlyValid = true;
            // Get the geometry rules for the current IC Type
            var geometryRules = _rules.ReturnIcGisTypeSettings(icType)?.GeometryRules;
            if (geometryRules == null)
            {
                shapeToValidate.IsValid = false;
                shapeToValidate.Status = "Missing Geometry Rules";
                recordShapeCheckFailure($"Could not find Geometry Rules for the IC Type '{icType}'.");
                return;
            }

            var geometry = shapeToValidate.Geometry;

            // 1. Check for and remove Z and M values.
            if (geometry.HasZ || geometry.HasM)
            {
                var repairTest = _namedTests.returnNewTestResult("GIS_Shape_Repair", shapeToValidate.SourceFile, IcTestResult.TestType.Shape);
                repairTest.Passed = true;

                // Use PolygonBuilderEx to create a new 2D polygon from the ZM coordinates.
                var flatPolygon = new PolygonBuilderEx(geometry.SpatialReference)
                {
                    HasZ = false,
                    HasM = false
                };
                flatPolygon.AddParts(geometry.Parts);

                geometry = flatPolygon.ToGeometry(); // Use the new flat polygon for all subsequent checks.
                shapeToValidate.Geometry = geometry as Polygon; // Update the shapeItem
                shapeToValidate.Status = "Repaired (ZM Removed)";
                repairTest.AddComment($"Shape with original ID {shapeToValidate.ShapeReferenceId} was a Polygon ZM and has been converted to a 2D Polygon.");
                parentTestResult.AddSubordinateTestResult(repairTest);
            }



            // 2: Check and Reproject Spatial Reference ---
            var requiredSr = SpatialReferenceBuilder.CreateSpatialReference(RequiredWkid);
            if (geometry.SpatialReference == null || !geometry.SpatialReference.IsEqual(requiredSr))
            {
                try
                {
                    // If not, reproject it.
                    var projectedGeometry = GeometryEngine.Instance.Project(geometry, requiredSr);
                    if (projectedGeometry != null)
                    {
                        shapeToValidate.Geometry = projectedGeometry as Polygon; // Update the geometry
                       // _log.RecordMessage($"Shape {shapeToValidate.ShapeReferenceId} was reprojected to NJ State Plane.", BisLogMessageType.Note);
                        geometry = shapeToValidate.Geometry;
                    }
                    else
                    {
                        shapeToValidate.IsValid = false;
                        shapeToValidate.Status = "Projection returned a null geometry";
                        recordShapeCheckFailure("Projection returned a null geometry.");
                        return;
                        //throw new Exception("Projection returned a null geometry.");
                    }
                }
                catch (Exception ex)
                {
                    _log.RecordError($"Failed to reproject geometry for shape {shapeToValidate.ShapeReferenceId}.", ex, "ValidateShape");
                    shapeToValidate.IsValid = false;
                    shapeToValidate.Status = "Reprojection Failed";
                    recordShapeCheckFailure($"The shape failed to reproject to the required coordinate system (NJ State Plane).");
                    return; // Stop validation if reprojection fails.
                }
            }

            // 3. Check if the shape is a polygon
            if (geometry.GeometryType != GeometryType.Polygon)
            {
                shapeToValidate.IsValid = false;
                shapeToValidate.Status = $"Invalid Type: {geometry.GeometryType}";
                recordShapeCheckFailure($"Incorrect geometry type. Expected Polygon, but found {geometry.GeometryType}.");
                return;
            }

            // 4. Check if the geometry is empty
            if (geometry.IsEmpty)
            {
                shapeToValidate.IsValid = false;
                shapeToValidate.Status = "Empty Geometry";
                recordShapeCheckFailure("The shape's geometry is empty.");
                return;
            }            

            // 5. Check for self-intersection and simplify if necessary (the modern way)
            // The GeometryEngine's SimplifyAsFeature method can fix many common geometry errors.
            if (!GeometryEngine.Instance.IsSimpleAsFeature(geometry))
            {
                var repairTest = _namedTests.returnNewTestResult("GIS_Shape_Repair", shapeToValidate.SourceFile, IcTestResult.TestType.Shape);
                repairTest.Passed = true;
                repairTest.AddComment($"Geometry error found in shape with original ID {shapeToValidate.ShapeReferenceId}. The geometry was automatically repaired.");
                parentTestResult.AddSubordinateTestResult(repairTest);
                geometry = (Polygon)GeometryEngine.Instance.SimplifyAsFeature(geometry, true); // true = allow endpoint changes
                shapeToValidate.Geometry = geometry as Polygon; // Update the geometry
                shapeToValidate.Status = "Repaired (Simplified)";
            }

            // 6. Check for inverted polygons. If area is negative, the orientation is inverted and reverse the orientation.
            double area = (geometry as Polygon).Area;
            if (area < 0)
            {
                var repairTest = _namedTests.returnNewTestResult("GIS_Shape_Repair", shapeToValidate.SourceFile, IcTestResult.TestType.Shape);
                repairTest.Passed = true;
                repairTest.AddComment($"Shape with original ID {shapeToValidate.ShapeReferenceId} had a negative area and its orientation was flipped.");
                parentTestResult.AddSubordinateTestResult(repairTest);

                // Flip the orientation and update the geometry
                geometry = (Polygon)GeometryEngine.Instance.ReverseOrientation(geometry);
                shapeToValidate.Geometry = geometry;
                shapeToValidate.Status = "Repaired (Flipped)";

                // Recalculate the area with the corrected orientation
                area = (geometry as Polygon).Area;
            }

            // 7. Check the area against the minimum threshold
            shapeToValidate.Area = area;
            if (Math.Abs(area) < geometryRules.Min_Area)
            {
                isShapeCurrentlyValid = false;
                //shapeToValidate.IsValid = false;
                shapeToValidate.Status = "Area Below Minimum";
                recordShapeCheckFailure("Area Below Minimum");
            }

            // 8.  Check if the shape is within the allowed extent ---
            var extent = geometry.Extent;
            if (extent.XMin < geometryRules.X_Min ||
                extent.XMax > geometryRules.X_Max ||
                extent.YMin < geometryRules.Y_Min ||
                extent.YMax > geometryRules.Y_Max)
            {
                //shapeToValidate.IsValid = false;
                isShapeCurrentlyValid = false;
                shapeToValidate.Status = "Outside Allowable Extent";
                recordShapeCheckFailure("Outside Allowable Extent");
            }

            // 9. Check the distance from the site, if a site location was provided.
            if (siteLocation != null && shapeToValidate.Geometry != null)
            {
                // Use the GeometryEngine to calculate the geodetic distance.
                double distance = GeometryEngine.Instance.Distance(shapeToValidate.Geometry, siteLocation);
                shapeToValidate.DistanceFromSite = distance; // Store the distance

                if (geometryRules != null && distance > geometryRules.SiteDistance)
                {
                    //shapeToValidate.IsValid = false;
                    isShapeCurrentlyValid = false;
                    shapeToValidate.Status = "Exceeds Max Distance from Site";
                    recordShapeCheckFailure("Outside Allowable Extent");
                }
            }

            // If all checks pass, the shape is considered valid
            shapeToValidate.IsValid = isShapeCurrentlyValid;
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
        private async Task<List<ShapeItem>> ReadFeaturesFromFileAsync(fileset fileset, string icType, IcTestResult parentTestResult)
        {
            var shapesInFile = new List<ShapeItem>();
            _log.RecordMessage($"Reading features from file: {fileset.fileName} (Type: {fileset.filesetType})", BIS_Log.BisLogMessageType.Note);

            try
            {
                switch (fileset.filesetType.ToLowerInvariant())
                {
                    case "shapefile":
                        string shpPath = Path.Combine(fileset.path, fileset.fileName + ".shp");
                        if (!File.Exists(shpPath))
                        {
                            _log.RecordError($"Shapefile not found at expected path: {shpPath}", null, "ReadFeaturesFromFileAsync");
                            fileset.validSet = false;
                            return shapesInFile;
                        }

                        await QueuedTask.Run(() =>
                        {
                            var shpConnectionPath = new FileSystemConnectionPath(new Uri(fileset.path), FileSystemDatastoreType.Shapefile);
                            using (var datastore = new FileSystemDatastore(shpConnectionPath))
                            using (var featureClass = datastore.OpenDataset<FeatureClass>(fileset.fileName))
                            {
                                // Call the single, unified helper method
                                shapesInFile.AddRange(ExtractShapesFromFeatureClass(featureClass, fileset, icType, parentTestResult));
                            }
                        });
                        break;

                    case "dwg":
                        string dwgPath = Path.Combine(fileset.path, fileset.fileName + ".dwg");
                        if (!File.Exists(dwgPath))
                        {
                            _log.RecordError($"DWG file not found at expected path: {dwgPath}", null, "ReadFeaturesFromFileAsync");
                            fileset.validSet = false;
                            return shapesInFile;
                        }

                        await QueuedTask.Run(() =>
                        {
                            var cadConnectionPath = new FileSystemConnectionPath(new Uri(fileset.path), FileSystemDatastoreType.Cad);
                            using (var cadDatastore = new FileSystemDatastore(cadConnectionPath))
                            {
                                // Process Polygons from the DWG
                                using (var polygonFC = cadDatastore.OpenDataset<FeatureClass>($"{fileset.fileName}.dwg:Polygon"))
                                {
                                    // Call the single, unified helper method
                                    shapesInFile.AddRange(ExtractShapesFromFeatureClass(polygonFC, fileset, icType, parentTestResult));
                                }
                                // Process Polylines from the DWG
                                using (var polylineFC = cadDatastore.OpenDataset<FeatureClass>($"{fileset.fileName}.dwg:Polyline"))
                                {
                                    // Call the single, unified helper method again for the polylines
                                    shapesInFile.AddRange(ExtractShapesFromFeatureClass(polylineFC, fileset, icType, parentTestResult));
                                }
                            }
                        });
                        break;

                    default:
                        _log.RecordMessage($"File type '{fileset.filesetType}' is not supported for feature extraction.", BisLogMessageType.Warning);
                        fileset.validSet = false;
                        break;
                }
            }
            catch (Exception ex)
            {
                string filePath = Path.Combine(fileset.path, fileset.fileName);
                _log.RecordError($"An unexpected error occurred reading features from '{filePath}'.", ex, "ReadFeaturesFromFileAsync");
                fileset.validSet = false;

                var unreadableTest = _namedTests.returnNewTestResult("GIS_FileReadable", fileset.fileName, IcTestResult.TestType.Submission);
                unreadableTest.Passed = false;
                unreadableTest.AddComment($"The dataset '{fileset.fileName}' could not be opened. It may be corrupt.");
                parentTestResult.AddSubordinateTestResult(unreadableTest);
            }

            if (!shapesInFile.Any())
            {
                _log.RecordMessage($"No processable polygon features were found in the file: {fileset.fileName}", BisLogMessageType.Warning);
            }

            return shapesInFile;
        }



        /// <summary>
        /// Extracts all polygon features from a given feature class and converts them to ShapeItem objects.
        /// </summary>
        private List<ShapeItem> ExtractShapesFromFeatureClass(FeatureClass featureClass, fileset sourceFileSet, string icType, IcTestResult parentTestResult)
        {
            var shapes = new List<ShapeItem>();
            if (featureClass == null) return shapes;

            var nameFilters = _rules.ReturnIcGisTypeSettings(icType)?.FeatureNameFilters;
            var fieldsToMine = _rules.ReturnIcGisTypeSettings(icType).FeatureFields.Where(f => f.DisplayInPreview).Select(f => f.Fieldname).ToList();
            var featureClassDef = featureClass.GetDefinition();

            // Determine if we are reading from a polyline layer (common in DWG files)
            bool isPolylineSource = featureClassDef.GetShapeType() == GeometryType.Polyline;

            using (var cursor = featureClass.Search(null, false))
            {
                while (cursor.MoveNext())
                {
                    if (cursor.Current is not Feature feature) continue;

                    Polygon polygon = null;
                    bool wasConvertedFromPolyline = false;

                    // --- START OF POLYLINE CONVERSION LOGIC ---
                    if (isPolylineSource)
                    {
                        if (feature.GetShape() is Polyline polyline && polyline.PointCount > 3)
                        {
                            var startPoint = polyline.Points.First();
                            var endPoint = polyline.Points.Last();

                            double closingDistance = GeometryEngine.Instance.Distance(startPoint, endPoint);
                            const double closingTolerance = 1.0; // 1 foot

                            if (closingDistance <= closingTolerance)
                            {
                                try
                                {
                                    // If the polyline is nearly closed but not exactly, we must rebuild it
                                    // to ensure it is perfectly closed before converting to a polygon.
                                    if (closingDistance > 0)
                                    {
                                        // 1. Get all the points from the original polyline.
                                        var allPoints = polyline.Points.ToList();
                                        // 2. Add the start point to the end of the list to close the loop.
                                        allPoints.Add(startPoint);
                                        // 3. Create a new, perfectly closed polyline from the updated point list.
                                        polyline = new PolylineBuilderEx(allPoints.Select(p => new Coordinate2D(p.X, p.Y)), polyline.SpatialReference).ToGeometry();
                                    }

                                    // Now convert the perfectly closed polyline to a polygon.
                                    polygon = new PolygonBuilderEx(polyline).ToGeometry();
                                    wasConvertedFromPolyline = true;
                                }
                                catch (Exception ex)
                                {
                                    _log.RecordError($"Failed to convert nearly-closed polyline with ID {feature.GetObjectID()} from {sourceFileSet.fileName}.", ex, "ExtractShapesFromFeatureClass");
                                    continue;
                                }
                            }
                        }
                    }
                    else
                    {
                        polygon = feature.GetShape() as Polygon;
                    }

                    if (polygon == null) continue; // If we couldn't get a valid polygon, skip to the next feature.

                    var shapeItem = new ShapeItem
                    {
                        Geometry = polygon,
                        SourceFile = sourceFileSet.fileName,
                        ShapeReferenceId = (int)feature.GetObjectID(),
                        IsValid = true, // It will be validated in the next step
                        Status = wasConvertedFromPolyline ? "Repaired (Converted)" : "Pending Validation",
                        ShapeType = wasConvertedFromPolyline ? "Polygon (from Polyline)" : "Polygon"
                    };

                    if (nameFilters != null)
                    {
                        bool descriptionSet = false;
                        foreach (var filter in nameFilters)
                        {
                            string fieldName = filter.Key;
                            int fieldIndex = feature.FindField(fieldName);
                            if (fieldIndex != -1)
                            {
                                string attributeValue = feature[fieldIndex]?.ToString() ?? "";
                                if (filter.Value.Contains(attributeValue, StringComparer.OrdinalIgnoreCase))
                                {
                                    shapeItem.IsAutoSelected = true;
                                    shapeItem.Description = attributeValue;
                                    descriptionSet = true;
                                    break;
                                }
                            }
                        }

                        if (!descriptionSet)
                        {
                            foreach (string fieldName in nameFilters.Keys)
                            {
                                int fieldIndex = feature.FindField(fieldName);
                                if (fieldIndex != -1)
                                {
                                    string attributeValue = feature[fieldIndex]?.ToString();
                                    if (!string.IsNullOrEmpty(attributeValue))
                                    {
                                        shapeItem.Description = attributeValue;
                                        break;
                                    }
                                }
                            }
                        }
                    }

                    foreach (string fieldName in fieldsToMine)
                    {
                        int fieldIndex = feature.FindField(fieldName);
                        if (fieldIndex != -1) { shapeItem.Attributes[fieldName] = feature[fieldName]; }
                    }

                    shapes.Add(shapeItem);
                }
            }
            return shapes;
        }

        //private void ConvertAndAddPolyline(Feature feature, Polyline polyline, fileset sourceFileSet, IcTestResult parentTestResult, List<ShapeItem> shapeList, bool isAutoSelected)
        //{
        //    if (polyline.PointCount > 0 && polyline.Points.First().IsEqual(polyline.Points.Last()))
        //    {
        //        try
        //        {
        //            var polygon = new PolygonBuilderEx(polyline).ToGeometry();
        //            var shapeItem = new ShapeItem
        //            {
        //                Geometry = polygon,
        //                SourceFile = sourceFileSet.fileName,
        //                ShapeReferenceId = (int)feature.GetObjectID(),
        //                ShapeType = "Polygon (from Polyline)",
        //                IsValid = true,
        //                Status = "Repaired (Converted)",
        //                IsAutoSelected = isAutoSelected
        //            };

        //            var repairTest = _namedTests.returnNewTestResult("GIS_Shape_Repair", sourceFileSet.fileName, IcTestResult.TestType.Shape);
        //            repairTest.Passed = true;
        //            repairTest.AddComment($"Closed polyline with ID {shapeItem.ShapeReferenceId} was converted to a polygon." + (isAutoSelected ? " It was auto-selected." : ""));
        //            parentTestResult.AddSubordinateTestResult(repairTest);

        //            shapeList.Add(shapeItem);
        //        }
        //        catch (Exception ex)
        //        {
        //            _log.RecordError($"Failed to convert closed polyline with ID {feature.GetObjectID()}.", ex, "ConvertAndAddPolyline");
        //        }
        //    }
        //}

        /// <summary>
        /// Builds a SQL WHERE clause from a dictionary of field names and a list of values to match.
        /// </summary>
        private string ConstructWhereClauseFromFilters(Dictionary<string, List<string>> nameFilters, FeatureClass featureClass)
        {
            if (nameFilters == null || !nameFilters.Any() || featureClass == null)
            {
                return null;
            }

            var validClauses = new List<string>();
            var fields = featureClass.GetDefinition().GetFields();

            // Iterate through the filters (e.g., Key="Layer", Value=["CEA boundary", "CEA boundary-GIS"])
            foreach (var filter in nameFilters)
            {
                string fieldName = filter.Key;
                List<string> valuesToMatch = filter.Value;

                // 1. Check if the field exists in the feature class and if there are values to match
                if (valuesToMatch.Any() && fields.Any(f => f.Name.Equals(fieldName, StringComparison.OrdinalIgnoreCase)))
                {
                    // 2. Format the values for a SQL 'IN' clause (e.g., "'VALUE1', 'VALUE2'")
                    //    This includes converting to uppercase and escaping single quotes.
                    string formattedValues = string.Join(",", valuesToMatch.Select(v => $"'{v.ToUpper().Replace("'", "''")}'"));

                    // 3. Build the clause for this field (e.g., "UPPER(Layer) IN ('CEA BOUNDARY', 'CEA BOUNDARY-GIS')")
                    validClauses.Add($"UPPER({fieldName}) IN ({formattedValues})");
                }
            }

            if (!validClauses.Any())
            {
                return null; // No matching fields were found.
            }

            // 4. Combine all valid clauses with " OR "
            return string.Join(" OR ", validClauses);
        }
    }
}