using IC_Loader_Pro.Models;
using IC_Rules_2025;
using System.Collections.Generic;

namespace IC_Loader_Pro.Models
{
    /// <summary>
    /// A simple data class to hold the complete results of processing a single email.
    /// </summary>
    public class EmailProcessingResult
    {
        public IcTestResult TestResult { get; set; }
        public AttachmentAnalysisResult AttachmentAnalysis { get; set; }
        public List<ShapeItem> ShapeItems { get; set; } = new List<ShapeItem>();      
    }
}