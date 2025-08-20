using IC_Rules_2025;
using System;
using System.Threading.Tasks;
using static BIS_Log;
using static IC_Loader_Pro.Module1;

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
    }
}