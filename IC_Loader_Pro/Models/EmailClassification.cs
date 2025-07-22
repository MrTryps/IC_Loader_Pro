using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using BIS_Tools_DataModels_2025;

namespace IC_Loader_Pro.Models
{


    /// <summary>
    /// Holds the results of an email classification operation.
    /// Based on the legacy OutlookTools.emailTypeResponse struct.
    /// </summary>
    public class EmailClassificationResult
    {
        public EmailType Type { get; set; } = EmailType.Unknown;
        public bool IsSubjectLineValid { get; set; } = true;
      
        /// <summary>
        /// A flag indicating that the user manually set the email type via the pop-up window.
        /// </summary>
        public bool WasManuallyClassified { get; set; } = false;
        // --------------------------

        public string InvalidReason { get; set; }
        public List<string> PrefIds { get; set; } = new List<string>();
        public List<string> AltIds { get; set; } = new List<string>();
        public List<string> ActivityNums { get; set; } = new List<string>();
        public string Note { get; set; }
    }
}
