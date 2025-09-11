using IC_Loader_Pro.Models;
using IC_Rules_2025;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using static BIS_Log;
using static IC_Loader_Pro.Module1;
using static IC_Rules_2025.IcTestResult;

namespace IC_Loader_Pro.Services
{
    public class TestResultService
    {
        /// <summary>
        /// Saves the entire hierarchy of a test result to the database,
        /// linking it to a specific deliverable ID.
        /// </summary>
        /// <param name="testResult">The root IcTestResult object to save.</param>
        /// <param name="deliverableId">The deliverable ID to associate the results with.</param>
        public async Task SaveTestResultsAsync(IcTestResult testResult, string deliverableId)
        {
            if (testResult == null) return;
            const string methodName = "SaveTestResultsAsync";

            try
            {
                // The IcTestResult class from your rules engine is expected to have a
                // method that handles its own database persistence. We run it on a
                // background thread to keep the UI responsive.
                await Task.Run(() => testResult.RecordResults(deliverableId));
                Log.RecordMessage($"Successfully saved test results for deliverable {deliverableId}.", BisLogMessageType.Note);
            }
            catch (Exception ex)
            {
                Log.RecordError($"Error saving test results for deliverable ID '{deliverableId}'.", ex, methodName);
                throw;
            }
        }

        /// <summary>
        /// Compiles all test results from the processing workflow into a single, final,
        /// hierarchical test result for saving and reporting.
        /// </summary>
        /// <param name="initialDeliverableResult">The root test result for the email/deliverable level.</param>
        /// <param name="approvedShapes">The final list of shapes the user approved for saving.</param>
        /// <param name="deliverableId">The new, permanent deliverable ID.</param>
        /// <param name="icType">The IC Type of the submission.</param>
        /// <param name="prefId">The final Pref ID for the submission.</param>
        /// <returns>A single, final IcTestResult object.</returns>
        public IcTestResult CompileFinalResults(
             IcTestResult initialDeliverableResult,
             ICollection<ShapeItem> approvedShapes,
             string deliverableId,
             string icType,
             string prefId)
        {
            var namedTests = new IcNamedTests(Log, PostGreTool);

            // 1. Create the final root test result that will contain everything.
            var finalResult = namedTests.returnNewTestResult("GIS_Deliverable_Root", deliverableId, TestType.Deliverable);

            // 2. Create a container for all the submission-level test results.
            var submissionContainerResult = namedTests.returnNewTestResult("GIS_CompiledSubResults", deliverableId, TestType.Deliverable);

            // 3. Find all the fileset-specific tests from the children of the initial result.
            if (initialDeliverableResult?.SubTestResults != null)
            {
                foreach (var childTest in initialDeliverableResult.SubTestResults)
                {
                    if (childTest.ResultType == TestType.Submission)
                    {
                        submissionContainerResult.AddSubordinateTestResult(childTest);
                    }
                }
            }

            // 4. Add the main results and the submission container to the final root.
            finalResult.AddSubordinateTestResult(initialDeliverableResult);
            finalResult.AddSubordinateTestResult(submissionContainerResult);

            // 5. Create a final test to record how many shapes the user approved.
            var shapesApprovedResult = namedTests.returnNewTestResult("GIS_ShapesPromoted", deliverableId, TestType.Deliverable);
            int approvedShapeCount = approvedShapes?.Count ?? 0;
            shapesApprovedResult.Passed = approvedShapeCount > 0;
            shapesApprovedResult.AddComment($"{approvedShapeCount} shapes approved to load.");
            finalResult.AddSubordinateTestResult(shapesApprovedResult);

            // 6. Add final parameters for use in the reply email.
            finalResult.addParameter("ShapeCount", approvedShapeCount.ToString());
            finalResult.addParameter("Ic_Type", icType);
            finalResult.addParameter("prefid", prefId);

            return finalResult;
        }
    }
}

