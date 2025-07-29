using IC_Loader_Pro.Models;
using IC_Rules_2025;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using static BIS_Log;

namespace IC_Loader_Pro.Services
{
    /// <summary>
    /// A service dedicated to processing email attachments: saving, unzipping, and identifying file sets.
    /// </summary>
    public class AttachmentService
    {
        private readonly BIS_Log _log;
        private readonly IC_Rules _rules;
        private readonly IcNamedTests _namedTests;
        private readonly BisFileTools _fileTool; // Assuming this service is available

        public AttachmentService(IC_Rules rules, IcNamedTests namedTests, BisFileTools fileTool, BIS_Log log)
        {
            _rules = rules;
            _namedTests = namedTests;
            _fileTool = fileTool;
            _log = log;
        }

        /// <summary>
        /// Saves all attachments from a given Outlook MailItem to a new, unique temporary folder.
        /// This version handles duplicate attachment filenames by appending a number.
        /// </summary>
        /// <param name="attachments">The collection of attachments from an Outlook.MailItem.</param>
        /// <returns>The full path to the new temporary directory containing the saved files.</returns>
        public string SaveAttachmentsToTempFolder(Microsoft.Office.Interop.Outlook.Attachments attachments)
        {
            if (attachments == null || attachments.Count == 0)
            {
                _log.RecordMessage("Email contains no attachments to save.", BisLogMessageType.Note);
                return null;
            }

            string tempFolderPath = Path.Combine(Path.GetTempPath(), "IC_Loader", Guid.NewGuid().ToString());
            Directory.CreateDirectory(tempFolderPath);
            _log.RecordMessage($"Created temporary attachment folder: {tempFolderPath}", BisLogMessageType.Note);

            var savedFilePaths = new List<string>();

            foreach (Microsoft.Office.Interop.Outlook.Attachment attachment in attachments)
            {
                try
                {
                    string sanitizedFileName = _fileTool.SanitizeFileName(attachment.FileName);
                    string fullPath = Path.Combine(tempFolderPath, sanitizedFileName);
                    string baseName = Path.GetFileNameWithoutExtension(sanitizedFileName);
                    string extension = Path.GetExtension(sanitizedFileName);
                    int count = 1;

                    // --- NEW: Loop to ensure filename is unique ---
                    while (File.Exists(fullPath))
                    {
                        string newFileName = $"{baseName} ({count++}){extension}";
                        fullPath = Path.Combine(tempFolderPath, newFileName);
                    }

                    attachment.SaveAsFile(fullPath);
                    savedFilePaths.Add(fullPath);
                }
                catch (Exception ex)
                {
                    _log.RecordError($"Failed to save attachment '{attachment.FileName}'.", ex, "SaveAttachmentsToTempFolder");
                }
            }

            _log.RecordMessage($"Successfully saved {savedFilePaths.Count} attachments.", BisLogMessageType.Note);
            return tempFolderPath;
        }

        public AttachmentAnalysisResult AnalyzeAttachments(string folderToSearch, string icType)
        {
            const string methodName = "AnalyzeAttachments";

            var analysisResult = new AttachmentAnalysisResult
            {
                TestResult = _namedTests.returnNewTestResult("GIS_Attachments_Tests_Passed", "", IcTestResult.TestType.Deliverable),
                TempFolderPath = folderToSearch
            };

            // --- THIS IS THE NEW LOGIC ---
            // First, check if there was even a folder created.
            // The TempFolderPath will only be set if attachments existed to be saved.
            if (string.IsNullOrEmpty(folderToSearch))
            {
                // This is not a code error, but a validation failure. The submission is invalid.
                analysisResult.TestResult.Passed = false;
                analysisResult.TestResult.AddComment("Email contains no attachments.");
                _log.RecordMessage("Attachment analysis determined the email has no attachments.", BisLogMessageType.Note);
                return analysisResult; // Exit immediately
            }

            try
            {
                // Step 1: Unzip any archive files.
                var unzipService = new UnzipService(_log);
                var unzipTestResult = _namedTests.returnNewTestResult("GIS_Attachments_Unzip_Passed", "", IcTestResult.TestType.Deliverable);

                // This call finds all .zip files and extracts them.
                var unzippedFilesInfo = unzipService.UnzipAllInDirectory(folderToSearch, deleteOriginalZip: true);

                if (unzippedFilesInfo.Any())
                {
                    unzipTestResult.Comments.Add($"Successfully extracted {unzippedFilesInfo.Count} zip file(s).");
                }
                else
                {
                    unzipTestResult.Comments.Add("No .zip files were found in the attachments.");
                }
                analysisResult.TestResult.AddSubordinateTestResult(unzipTestResult);


                // Step 2: Identify logical GIS filesets from the entire folder content.
                analysisResult.IdentifiedFileSets = _rules.ReturnFileSetsFromDirectory(folderToSearch, icType);


                // Step 3: Create a comprehensive list of all individual files.
                var allFilesFound = _fileTool.ListOfFilesInFolder(folderToSearch); // true = recursive

                foreach (string filePath in allFilesFound)
                {
                    // Find which zip file this file came from, if any.
                    var parentZipInfo = unzippedFilesInfo
                        .FirstOrDefault(zipInfo => filePath.StartsWith(zipInfo.ExtractionPath, StringComparison.OrdinalIgnoreCase));

                    analysisResult.AllFiles.Add(new AnalyzedFile
                    {
                        FileName = Path.GetFileName(filePath),
                        CurrentPath = Path.GetDirectoryName(filePath),
                        // If the file was in a zip, record the zip's name as its original path.
                        OriginalPath = parentZipInfo.OriginalZipFileName ?? string.Empty
                    });
                }
            }
            catch (Exception ex)
            {
                _log.RecordError("An error occurred during attachment analysis.", ex, methodName);
                analysisResult.TestResult.Passed = false;
                analysisResult.TestResult.Comments.Add($"Fatal error during attachment analysis: {ex.Message}");
            }

            return analysisResult;
        }
               
    }
}