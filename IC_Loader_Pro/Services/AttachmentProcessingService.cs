using BIS_Tools_DataModels_2025;
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

            var analysisResult = new AttachmentAnalysisResult();
            IcTestResult rootTest = null;

            // Wrap the creation of the root test result in its own try-catch block.
            try
            {
                rootTest = _namedTests.returnNewTestResult("GIS_Attachments_Tests_Passed", "", IcTestResult.TestType.Deliverable);
            }
            catch (Exception ex)
            {
                _log.RecordError("Could not create the root attachment test result. A default will be used.", ex, methodName);
                // If it fails, create a default, failing test result so the process can continue.
                var tempRule = new IcTestRule { Name = "AttachmentProcessingError", Action = TestActionResponse.Fail };
                rootTest = new IcTestResult(tempRule, "", IcTestResult.TestType.Deliverable, _log, null, _namedTests)
                {
                    Passed = false,
                    Comments = { "Critical error: Could not load the 'GIS_Attachments_Tests_Passed' rule from the database." }
                };
            }
            analysisResult.TestResult = rootTest;
            analysisResult.TempFolderPath = folderToSearch;


            // First, check if there was even a folder created.
            if (string.IsNullOrEmpty(folderToSearch))
            {
                analysisResult.TestResult.Passed = false;
                analysisResult.TestResult.AddComment("Email contains no attachments.");
                _log.RecordMessage("Attachment analysis determined the email has no attachments.", BisLogMessageType.Note);
                return analysisResult;
            }

            try
            {
                var fileNames = Directory.GetFiles(folderToSearch).Select(Path.GetFileName).ToList();
                var duplicateFiles = fileNames.GroupBy(f => f, StringComparer.OrdinalIgnoreCase)
                                              .Where(g => g.Count() > 1)
                                              .Select(g => g.Key)
                                              .ToList();

                if (duplicateFiles.Any())
                {
                    var duplicateTest = _namedTests.returnNewTestResult("GIS_DuplicateFilenamesInAttachments", "", IcTestResult.TestType.Deliverable);
                    duplicateTest.Passed = false;
                    duplicateTest.AddComment($"Duplicate filenames found: {string.Join(", ", duplicateFiles)}");
                    analysisResult.TestResult.AddSubordinateTestResult(duplicateTest);
                    analysisResult.TestResult.Passed = false;
                    return analysisResult; // Stop processing immediately
                }
            }
            catch (Exception ex)
            {
                _log.RecordError("An error occurred during the duplicate filename check.", ex, methodName);
                analysisResult.TestResult.Passed = false;
                analysisResult.TestResult.AddComment("A critical error occurred while checking for duplicate filenames.");
                return analysisResult;
            }
            // --- END OF NEW DUPLICATE FILENAME CHECK ---

            try
            {
                // Step 1: Unzip any archive files.
                var unzipService = new UnzipService(_log);
                var unzippedFilesInfo = unzipService.UnzipAllInDirectory(folderToSearch, deleteOriginalZip: true);

                IcTestResult unzipTestResult = null;
                // Wrap the creation of the unzip test result in its own try-catch block.
                try
                {
                    unzipTestResult = _namedTests.returnNewTestResult("GIS_Attachments_Unzip_Passed", "", IcTestResult.TestType.Deliverable);
                    if (unzippedFilesInfo.Any())
                    {
                        unzipTestResult.Comments.Add($"Successfully extracted {unzippedFilesInfo.Count} zip file(s).");
                    }
                    else
                    {
                        unzipTestResult.Comments.Add("No .zip files were found in the attachments.");
                    }
                    analysisResult.TestResult.AddSubordinateTestResult(unzipTestResult);
                }
                catch (Exception ex)
                {
                    _log.RecordError("Could not create the unzip test result. This step will be skipped in the results.", ex, methodName);
                }

                // Step 2: Identify logical GIS filesets by searching the root folder and all unzipped sub-folders.
                var allIdentifiedFileSets = new List<BIS_Tools_DataModels_2025.fileset>();
                allIdentifiedFileSets.AddRange(_rules.ReturnFileSetsFromDirectory(folderToSearch, icType, false));
                foreach (var unzippedInfo in unzippedFilesInfo)
                {
                    allIdentifiedFileSets.AddRange(_rules.ReturnFileSetsFromDirectory(unzippedInfo.ExtractionPath, icType, false));
                }
                analysisResult.IdentifiedFileSets = allIdentifiedFileSets;

                foreach (var fileset in analysisResult.IdentifiedFileSets.Where(fs => !fs.validFileSet))
                {
                    var incompleteTest = _namedTests.returnNewTestResult("GIS_Incomplete_Dataset", fileset.fileName, IcTestResult.TestType.Submission);
                    incompleteTest.Passed = false;
                    incompleteTest.AddComment($"The dataset '{fileset.fileName}' is incomplete or missing required files (e.g., .dbf, .shx).");
                    analysisResult.TestResult.AddSubordinateTestResult(incompleteTest);
                    analysisResult.TestResult.Passed = false;
                }

                // Step 3: Create a comprehensive list of all individual files.
                var allFilesFound = _fileTool.ListOfFilesInFolder(folderToSearch);

                foreach (string filePath in allFilesFound)
                {
                    var parentZipInfo = unzippedFilesInfo
                        .FirstOrDefault(zipInfo => filePath.StartsWith(zipInfo.ExtractionPath, StringComparison.OrdinalIgnoreCase));

                    analysisResult.AllFiles.Add(new AnalyzedFile
                    {
                        FileName = Path.GetFileName(filePath),
                        CurrentPath = Path.GetDirectoryName(filePath),
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
        //public AttachmentAnalysisResult AnalyzeAttachments(string folderToSearch, string icType)
        //{
        //    const string methodName = "AnalyzeAttachments";

        //    var analysisResult = new AttachmentAnalysisResult
        //    {
        //        TestResult = _namedTests.returnNewTestResult("GIS_Attachments_Tests_Passed", "", IcTestResult.TestType.Deliverable),
        //        TempFolderPath = folderToSearch
        //    };

        //    // First, check if there was even a folder created.
        //    // The TempFolderPath will only be set if attachments existed to be saved.
        //    if (string.IsNullOrEmpty(folderToSearch))
        //    {
        //        // This is not a code error, but a validation failure. The submission is invalid.
        //        analysisResult.TestResult.Passed = false;
        //        analysisResult.TestResult.AddComment("Email contains no attachments.");
        //        _log.RecordMessage("Attachment analysis determined the email has no attachments.", BisLogMessageType.Note);
        //        return analysisResult; // Exit immediately
        //    }

        //    try
        //    {
        //        // Step 1: Unzip any archive files.
        //        var unzipService = new UnzipService(_log);
        //        var unzipTestResult = _namedTests.returnNewTestResult("GIS_Attachments_Unzip_Passed", "", IcTestResult.TestType.Deliverable);

        //        // This call finds all .zip files and extracts them.
        //        var unzipResult = unzipService.UnzipAllInDirectory(folderToSearch, deleteOriginalZip: true);
        //        var unzippedFilesInfo = unzipResult.Succeeded;

        //        if (unzipResult.FailedFiles.Any())
        //        {
        //            // If any files failed, the test fails.
        //            unzipTestResult.Passed = false;
        //            unzipTestResult.AddComment($"Failed to extract {unzipResult.FailedFiles.Count} zip file(s): {string.Join(", ", unzipResult.FailedFiles)}. They may be corrupt.");
        //        }
        //        else if (unzipResult.Succeeded.Any())
        //        {
        //            unzipTestResult.Comments.Add($"Successfully extracted {unzipResult.Succeeded.Count} zip file(s).");
        //        }
        //        else
        //        {
        //            unzipTestResult.Comments.Add("No .zip files were found in the attachments.");
        //        }
        //        analysisResult.TestResult.AddSubordinateTestResult(unzipTestResult);


        //        // Step 2: Identify logical GIS filesets from the entire folder content.
        //        analysisResult.IdentifiedFileSets = _rules.ReturnFileSetsFromDirectory(folderToSearch, icType);


        //        foreach (var fileset in analysisResult.IdentifiedFileSets.Where(fs => !fs.validFileSet))
        //        {
        //            var incompleteTest = _namedTests.returnNewTestResult("GIS_Incomplete_Dataset", fileset.fileName, IcTestResult.TestType.Submission);
        //            incompleteTest.Passed = false; // This is a failing test.
        //            incompleteTest.AddComment($"The dataset '{fileset.fileName}' is incomplete or missing required files (e.g., .dbf, .shx).");
        //            analysisResult.TestResult.AddSubordinateTestResult(incompleteTest);
        //            //analysisResult.TestResult.Passed = false; // Mark the parent attachment test as failed.
        //        }


        //        // Step 3: Create a comprehensive list of all individual files.
        //        var allFilesFound = _fileTool.ListOfFilesInFolder(folderToSearch); // true = recursive

        //        foreach (string filePath in allFilesFound)
        //        {
        //            // Find which zip file this file came from, if any.
        //            var parentZipInfo = unzippedFilesInfo
        //                .FirstOrDefault(zipInfo => filePath.StartsWith(zipInfo.ExtractionPath, StringComparison.OrdinalIgnoreCase));

        //            analysisResult.AllFiles.Add(new AnalyzedFile
        //            {
        //                FileName = Path.GetFileName(filePath),
        //                CurrentPath = Path.GetDirectoryName(filePath),
        //                // If the file was in a zip, record the zip's name as its original path.
        //                OriginalPath = parentZipInfo.OriginalZipFileName ?? string.Empty
        //            });
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        _log.RecordError("An error occurred during attachment analysis.", ex, methodName);
        //        analysisResult.TestResult.Passed = false;
        //        analysisResult.TestResult.Comments.Add($"Fatal error during attachment analysis: {ex.Message}");
        //    }

        //    return analysisResult;
        //}

    }
}