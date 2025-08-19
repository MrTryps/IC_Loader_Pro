// In IC_Loader_Pro/Services/NotesService.cs

using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using static IC_Loader_Pro.Module1;

namespace IC_Loader_Pro.Services
{
    public class NotesService
    {
        /// <summary>
        /// Records a general comment in the database, associated with a specific reference ID.
        /// </summary>
        /// <param name="referenceId">The ID of the record to associate the comment with.</param>
        /// <param name="commentText">The text of the comment to record.</param>
        /// <param name="commentType">Optional. The type of comment to record. Defaults to "comment".</param>
        public async Task RecordNoteAsync(string referenceId, string commentText)
        {
            const string methodName = "RecordNoteAsync";

            // The parameter names here should match the names in your database function.
            var paramDict = new Dictionary<string, object>
            {
                { "ref_id", referenceId },
                { "msg", commentText }
               
            };

            try
            {
                await Task.Run(() =>
                    PostGreTool.ExecuteNamedQuery("recordGisComment", paramDict)
                );
            }
            catch (Exception ex)
            {
                Log.RecordError($"Error recording note for Reference ID '{referenceId}'.", ex, methodName);
                throw;
            }
        }
    }
}