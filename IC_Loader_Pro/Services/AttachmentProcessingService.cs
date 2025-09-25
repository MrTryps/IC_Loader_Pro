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


        public AttachmentAnalysisResult AnalyzeAttachments(string folderToSearch, string icType, List<EmailItem.AttachmentItem> attachments)
        {
            const string methodName = "AnalyzeAttachments";

            var analysisResult = new AttachmentAnalysisResult
            {
                TestResult = _namedTests.returnNewTestResult("GIS_Attachments_Tests", "", IcTestResult.TestType.Deliverable),
                TempFolderPath = folderToSearch
            };

            if (string.IsNullOrEmpty(folderToSearch))
            {
                analysisResult.TestResult.Passed = false;
                analysisResult.TestResult.AddComment("Email contains no attachments.");
                return analysisResult;
            }

            try
            {
                // 1. Perform the duplicate filename check FIRST.
                var duplicateOriginalFilenames = attachments
                                                   .GroupBy(a => a.OriginalFileName, StringComparer.OrdinalIgnoreCase)
                                                   .Where(g => g.Count() > 1)
                                                   .Select(g => g.Key)
                                                   .ToList();

                var problematicBaseNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

                if (duplicateOriginalFilenames.Any())
                {
                    var multiFileDuplicates = new List<string>();
                    foreach (var dupName in duplicateOriginalFilenames)
                    {
                        var rule = _rules.ReturnFilesetRuleForExtension(Path.GetExtension(dupName).TrimStart('.'));
                        if (rule != null && rule.RequiredExtensions.Count > 1)
                        {
                            multiFileDuplicates.Add(dupName);
                            problematicBaseNames.Add(Path.GetFileNameWithoutExtension(dupName));
                        }
                    }

                    if (multiFileDuplicates.Any())
                    {
                        var duplicateTest = _namedTests.returnNewTestResult("GIS_DuplicateFilenamesInAttachments", "", IcTestResult.TestType.Submission);
                        duplicateTest.Passed = false;
                        duplicateTest.AddComment($"The submission contains multiple multi-file datasets with the same filename(s): {string.Join(", ", multiFileDuplicates.Distinct())}. These files will be ignored.");
                        analysisResult.TestResult.AddSubordinateTestResult(duplicateTest);
                    }
                }

                // 2. Unzip any archive files.
                var unzipService = new UnzipService(_log);
                var unzippedFilesInfo = unzipService.UnzipAllInDirectory(folderToSearch, deleteOriginalZip: true);
                var unzipTestResult = _namedTests.returnNewTestResult("GIS_Attachments_Unzip", "", IcTestResult.TestType.Deliverable);
                unzipTestResult.AddComment(unzippedFilesInfo.Any() ? $"Successfully extracted {unzippedFilesInfo.Count} zip file(s)." : "No .zip files were found in the attachments.");
                analysisResult.TestResult.AddSubordinateTestResult(unzipTestResult);

                var allIdentifiedFileSets = _rules.ReturnFileSetsFromDirectory_NewMethod(folderToSearch, icType, true);

                // 3. Filter out any filesets that were identified as having duplicate names earlier.
                analysisResult.IdentifiedFileSets = allIdentifiedFileSets
                                                     .Where(fs => !problematicBaseNames.Contains(fs.fileName))
                                                     .ToList();



                //    // 3. Create a list of all folders to search, which includes the root folder AND all new subfolders from the unzipping process.
                //    var foldersToSearch = new List<string> { folderToSearch };
                //    foldersToSearch.AddRange(unzippedFilesInfo.Select(info => info.ExtractionPath));

                //    var allIdentifiedFileSets = new List<fileset>();

                //    // 4. Loop through each folder and find filesets within it (non-recursively).
                //    foreach (var currentFolder in foldersToSearch)
                //    {
                //        allIdentifiedFileSets.AddRange(_rules.ReturnFileSetsFromDirectory_NewMethod(currentFolder, icType, true));
                //    }

                //    var uniqueFilesets = allIdentifiedFileSets
                //.GroupBy(fs => Path.Combine(fs.path, fs.fileName))
                //.Select(g => g.First())
                //.ToList();


                //    // 5. Filter out any filesets that were identified as having duplicate names earlier.
                //    analysisResult.IdentifiedFileSets = allIdentifiedFileSets
                //                                         .Where(fs => !problematicBaseNames.Contains(fs.fileName))
                //                                         .ToList();

                // 6. Check for incomplete filesets among the valid ones.
                foreach (var fileset in analysisResult.IdentifiedFileSets.Where(fs => !fs.validSet))
                {
                    var incompleteTest = _namedTests.returnNewTestResult("GIS_Incomplete_Dataset", fileset.fileName, IcTestResult.TestType.Submission);
                    incompleteTest.Passed = false;
                    incompleteTest.addParameter("filename", fileset.fileName);
                    analysisResult.TestResult.AddSubordinateTestResult(incompleteTest);
                    analysisResult.TestResult.Passed = false;
                }

                // 7. Create a comprehensive list of all individual files.
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

    }
}