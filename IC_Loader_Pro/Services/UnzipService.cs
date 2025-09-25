using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;

namespace IC_Loader_Pro.Services
{
    public class UnzipService
    {
        private readonly BIS_Log _log;

        public struct UnzippedFileInfo
        {
            public string OriginalZipFileName { get; set; }
            public string ExtractionPath { get; set; }
        }

        public UnzipService(BIS_Log log)
        {
            _log = log;
        }

        /// <summary>
        /// Finds and extracts all zip files within a given directory, including nested zips,
        /// until no more zip files are found.
        /// </summary>
        public List<UnzippedFileInfo> UnzipAllInDirectory(string directoryToSearch, bool deleteOriginalZip = false)
        {
            var allUnzippedFiles = new List<UnzippedFileInfo>();
            if (string.IsNullOrEmpty(directoryToSearch) || !Directory.Exists(directoryToSearch))
            {
                _log.RecordError($"The directory provided for unzipping does not exist: '{directoryToSearch}'.", null, nameof(UnzipAllInDirectory));
                return allUnzippedFiles;
            }

            // Keep looping as long as we are finding and extracting new zip files.
            while (true)
            {
                // Find all zip files in the root directory AND all subdirectories.
                var zipFiles = Directory.GetFiles(directoryToSearch, "*.zip", SearchOption.AllDirectories);

                // If no zip files are found in this pass, we are done.
                if (!zipFiles.Any())
                {
                    break;
                }

                // This list will hold the results from the current pass.
                var newlyUnzippedInThisPass = new List<UnzippedFileInfo>();

                foreach (var zipFile in zipFiles)
                {
                    try
                    {
                        string destinationFolder = Path.Combine(
                            Path.GetDirectoryName(zipFile),
                            Path.GetFileNameWithoutExtension(zipFile)
                        );

                        // Ensure the destination folder name is unique.
                        int count = 1;
                        string originalDestination = destinationFolder;
                        while (Directory.Exists(destinationFolder))
                        {
                            destinationFolder = $"{originalDestination}_{count++}";
                        }

                        Directory.CreateDirectory(destinationFolder);
                        ZipFile.ExtractToDirectory(zipFile, destinationFolder);

                        _log.RecordMessage($"Successfully extracted '{zipFile}' to '{destinationFolder}'.", BIS_Log.BisLogMessageType.Note);

                        newlyUnzippedInThisPass.Add(new UnzippedFileInfo
                        {
                            OriginalZipFileName = zipFile,
                            ExtractionPath = destinationFolder
                        });

                        if (deleteOriginalZip)
                        {
                            File.Delete(zipFile);
                        }
                    }
                    catch (InvalidDataException ex)
                    {
                        _log.RecordError($"The file '{zipFile}' is not a valid zip archive and was skipped.", ex, nameof(UnzipAllInDirectory));
                        // If it's not a valid zip, delete it so we don't try it again in the next loop.
                        if (deleteOriginalZip) File.Delete(zipFile);
                    }
                    catch (Exception ex)
                    {
                        _log.RecordError($"An unexpected error occurred while extracting '{zipFile}'. It was skipped.", ex, nameof(UnzipAllInDirectory));
                    }
                }

                // If we didn't successfully unzip anything in this pass, break the loop to prevent an infinite loop on a corrupt file.
                if (!newlyUnzippedInThisPass.Any())
                {
                    break;
                }

                allUnzippedFiles.AddRange(newlyUnzippedInThisPass);
            }

            return allUnzippedFiles;
        }
    }
}