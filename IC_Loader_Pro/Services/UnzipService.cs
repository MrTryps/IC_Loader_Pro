using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using static IC_Loader_Pro.Services.UnzipService;

namespace IC_Loader_Pro.Services
{
    /// <summary>
    /// A result class to hold both successes and failures from the unzip process.
    /// </summary>
    public class UnzipResult
    {
        public List<UnzippedFileInfo> Succeeded { get; } = new List<UnzippedFileInfo>();
        public List<string> FailedFiles { get; } = new List<string>();
    }

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
        /// Finds and extracts all zip files within a given directory, returning a result
        /// object that contains lists of both successful and failed extractions.
        /// </summary>
        public List<UnzippedFileInfo> UnzipAllInDirectory(string directoryToSearch, bool deleteOriginalZip = false)
        {
            var unzippedFiles = new List<UnzippedFileInfo>();

            if (string.IsNullOrEmpty(directoryToSearch) || !Directory.Exists(directoryToSearch))
            {
                _log.RecordError($"The directory provided for unzipping does not exist: '{directoryToSearch}'.", null, nameof(UnzipAllInDirectory));
                return unzippedFiles; // Return an empty list
            }

            // Get all files with a .zip extension in the directory.
            // SearchOption.AllDirectories makes this recursive automatically.
            var zipFiles = Directory.GetFiles(directoryToSearch, "*.zip", SearchOption.AllDirectories);

            foreach (var zipFile in zipFiles)
            {
                try
                {
                    // Create a unique destination folder for the contents of each zip file.
                    // This prevents files from different zips from overwriting each other.
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

                    // --- The Core Logic ---
                    // This one line handles the entire extraction process.
                    ZipFile.ExtractToDirectory(zipFile, destinationFolder);

                    _log.RecordMessage($"Successfully extracted '{zipFile}' to '{destinationFolder}'.", BIS_Log.BisLogMessageType.Note);

                    unzippedFiles.Add(new UnzippedFileInfo
                    {
                        OriginalZipFileName = zipFile,
                        ExtractionPath = destinationFolder
                    });

                    // Optionally, delete the original .zip file after extraction.
                    if (deleteOriginalZip)
                    {
                        File.Delete(zipFile);
                    }
                }
                catch (InvalidDataException ex)
                {
                    _log.RecordError($"The file '{zipFile}' is not a valid zip archive and could not be extracted.", ex, nameof(UnzipAllInDirectory));
                    // Continue to the next file
                }
                catch (Exception ex)
                {
                    _log.RecordError($"An unexpected error occurred while extracting '{zipFile}'.", ex, nameof(UnzipAllInDirectory));
                    // Continue to the next file
                }
            }

            return unzippedFiles;
        }
    }
}