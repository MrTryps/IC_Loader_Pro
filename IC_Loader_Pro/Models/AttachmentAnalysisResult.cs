using BIS_Tools_DataModels_2025;
using IC_Rules_2025;
using System.Collections.Generic;

namespace IC_Loader_Pro.Models
{
    /// <summary>
    /// Holds the results of processing and analyzing a folder of email attachments.
    /// Replaces the legacy GISFileSearchResults structure.
    /// </summary>
    public class AttachmentAnalysisResult
    {
        /// <summary>
        /// A list of logical filesets (e.g., shapefiles) identified from the attachments.
        /// </summary>
        public List<Fileset> IdentifiedFileSets { get; set; } = new List<Fileset>();

        /// <summary>
        /// A list of every individual file found after extraction.
        /// </summary>
        public List<AnalyzedFile> AllFiles { get; set; } = new List<AnalyzedFile>();

        /// <summary>
        /// The master test result for the attachment processing phase.
        /// </summary>
        public IcTestResult TestResult { get; set; }

        /// <summary>
        /// The root temporary directory where the attached files are saved.
        /// </summary>
        public string TempFolderPath { get; set; }
    }

    /// <summary>
    /// Represents a single physical file found during attachment processing.
    /// Replaces the legacy fileInfo structure.
    /// </summary>
    public class AnalyzedFile
    {
        public string FileName { get; set; }
        public string OriginalPath { get; set; }
        public string CurrentPath { get; set; }
    }
}