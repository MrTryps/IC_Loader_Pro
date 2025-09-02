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
        public UnzipResult UnzipAllInDirectory(string directoryToSearch, bool deleteOriginalZip = false)
        {
            var result = new UnzipResult();

            if (string.IsNullOrEmpty(directoryToSearch) || !Directory.Exists(directoryToSearch))
            {
                _log.RecordError($"The directory provided for unzipping does not exist: '{directoryToSearch}'.", null, nameof(UnzipAllInDirectory));
                return result;
            }

            var zipFiles = Directory.GetFiles(directoryToSearch, "*.zip", SearchOption.AllDirectories);

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

                    result.Succeeded.Add(new UnzippedFileInfo
                    {
                        OriginalZipFileName = zipFile,
                        ExtractionPath = destinationFolder
                    });

                    if (deleteOriginalZip)
                    {
                        File.Delete(zipFile);
                    }
                }
                catch (Exception ex)
                {
                    _log.RecordError($"Failed to extract '{zipFile}'. It may be corrupt or invalid.", ex, nameof(UnzipAllInDirectory));
                    // Add the failed file to our result list
                    result.FailedFiles.Add(Path.GetFileName(zipFile));
                }
            }
            return result;
        }
    }
}