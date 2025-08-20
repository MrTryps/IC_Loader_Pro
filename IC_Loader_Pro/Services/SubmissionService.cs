using BIS_Tools_DataModels_2025;
using IC_Loader_Pro.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Threading.Tasks;
using static BIS_Log;
using static IC_Loader_Pro.Module1;

namespace IC_Loader_Pro.Services
{
    public class SubmissionService
    {
        /// <summary>
        /// Creates a submission record in the database for each fileset.
        /// </summary>
        /// <returns>A dictionary mapping the original fileName to its new Submission ID.</returns>
        public async Task<Dictionary<string, string>> RecordSubmissionsAsync(string deliverableId, string icType, List<fileset> filesets)
        {
            var submissionIdMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            if (filesets == null) return submissionIdMap;

            foreach (var fs in filesets)
            {
                var paramDict = new Dictionary<string, object>
                {
                    { "DelId", deliverableId },
                    { "IcType", icType },
                    { "SOURCE_FILETYPE", fs.filesetType },
                    { "fileName", fs.fileName },
                    { "filePath", fs.path },
                    { "LogPath", Log.FileNameAndPath },
                    { "DSValidity", fs.validSet ? "Valid" : "Incomplete" }
                };

                string subId = await Task.Run(() => PostGreTool.ExecuteNamedQuery("NEWGISSUBMISSION", paramDict) as string);
                Log.RecordMessage($"New submission ID ({subId}) created for fileset '{fs.fileName}'.", BisLogMessageType.Note);
                await UpdateSubmissionStatusAsync(subId, "Recorded");
                submissionIdMap[fs.fileName] = subId;
            }
            return submissionIdMap;
        }

        /// <summary>
        /// Records all individual physical files associated with the deliverable.
        /// </summary>
        public async Task RecordPhysicalFilesAsync(string deliverableId, List<AnalyzedFile> allFiles, Dictionary<string, string> submissionIdMap)
        {
            if (allFiles == null) return;

            foreach (var file in allFiles)
            {
                string fileNameWithoutExt = Path.GetFileNameWithoutExtension(file.FileName);
                string submissionId = submissionIdMap.GetValueOrDefault(fileNameWithoutExt);

                var columns = new Dictionary<string, object>
                {
                    { "DELIVERABLE_ID", deliverableId },
                    { "CURRENT_FILENAME", file.FileName },
                    { "CURRENT_PATH", file.CurrentPath },
                    { "ORIGINAL_FILENAME", file.FileName },
                    { "ORIGINAL_PATH", file.OriginalPath }
                };

                if (!string.IsNullOrEmpty(submissionId))
                {
                    columns.Add("SUBMISSION_ID", submissionId);
                    string fullPath = Path.Combine(file.CurrentPath, file.FileName);
                    columns.Add("MD5_HASH", await GetFileMd5HashAsync(fullPath));
                }

                var fieldList = string.Join(", ", columns.Keys);
                var valuePlaceholders = string.Join(", ", columns.Keys.Select(k => "?"));
                var sql = $"INSERT INTO srp_gis_deliverable_files ({fieldList}) VALUES ({valuePlaceholders})";

                await Task.Run(() => PostGreTool.ExecuteRawQuery(sql, columns.Values.ToList(), "NOTHING"));
            }
        }

        /// <summary>
        /// Moves all files for the processed submissions to a final network location.
        /// </summary>
        public async Task MoveAllSubmissionsAsync(List<fileset> filesets, Dictionary<string, string> submissionIdMap, string asSubmittedRootPath)
        {
            if (filesets == null) return;

            foreach (var fs in filesets)
            {
                if (submissionIdMap.TryGetValue(fs.fileName, out string submissionId))
                {
                    string destinationFolder = Path.Combine(asSubmittedRootPath, submissionId);
                    Directory.CreateDirectory(destinationFolder);

                    foreach (var ext in fs.extensions)
                    {
                        string sourceFile = Path.Combine(fs.path, $"{fs.fileName}.{ext}");
                        string destFile = Path.Combine(destinationFolder, $"{fs.fileName}.{ext}");
                        if (File.Exists(sourceFile))
                        {
                            File.Copy(sourceFile, destFile, true);
                        }
                    }
                    Log.RecordMessage($"Moved files for submission '{submissionId}' to '{destinationFolder}'.", BisLogMessageType.Note);

                    var paramDict = new Dictionary<string, object> { { "Path", destinationFolder }, { "SubId", submissionId } };
                    await Task.Run(() => PostGreTool.ExecuteNamedQuery("Update_GIS_Sub_Path", paramDict));
                    await UpdateSubmissionStatusAsync(submissionId, "Moved to Network");
                }
            }
        }

        public async Task UpdateSubmissionCountsAsync(string submissionId, int goodCount, int duplicateCount)
        {
            const string methodName = "UpdateSubmissionCountsAsync";

            // Build the SQL UPDATE statement to set both count columns.
            string sql = "UPDATE srp_gis_submission SET good_record_count = ?, dup_record_count = ? WHERE submission_id = ?";

            // Build the list of parameters. The order must match the '?' placeholders in the SQL.
            var parameters = new List<object>
            {
                goodCount,
                duplicateCount,
                submissionId
            };

            try
            {
                await Task.Run(() => PostGreTool.ExecuteRawQuery(sql, parameters, "NOTHING"));
                Log.RecordMessage($"Updated counts for submission {submissionId}: Good={goodCount}, Dups={duplicateCount}.", BisLogMessageType.Note);
            }
            catch (Exception ex)
            {
                Log.RecordError($"Error updating counts for Submission ID '{submissionId}'.", ex, methodName);
                throw;
            }
        }

        private async Task UpdateSubmissionStatusAsync(string submissionId, string status)
        {
            var paramDict = new Dictionary<string, object>
            {
                { "SubId", submissionId },
                { "Status", status }
            };
            await Task.Run(() => PostGreTool.ExecuteNamedQuery("Update_GIS_Sub_Status", paramDict));
        }

        private async Task<string> GetFileMd5HashAsync(string filePath)
        {
            if (!File.Exists(filePath)) return null;

            using (var md5 = MD5.Create())
            {
                using (var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.Read))
                {
                    var hashBytes = await md5.ComputeHashAsync(stream);
                    return BitConverter.ToString(hashBytes).Replace("-", "").ToLowerInvariant();
                }
            }
        }
    }
}