using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IC_Loader_Pro.Models
{
    /// <summary>
    /// Defines the various types an incoming email can be classified as.
    /// Based on the legacy OutlookTools.emailType enum.
    /// </summary>
    public enum EmailType
    {
        Unknown = 0,
        Spam = 1,
        AutoResponse = 2,
        BlockedEmail = 3,
        Edd_New = 4,
        Edd_Resubmit = 5,
        EDD_Portal = 6,
        EDD_Fix = 7,
        CEA = 8,
        DNA = 9,
        IEC = 10,
        CKE = 11,
        WRS = 12,
        SRP_Forms = 13,
        SRP_Notifications = 14,
        EmptySubjectline = 15,
        Multiple = 98,
        Skip = 99
    }

    /// <summary>
    /// Holds the results of an email classification operation.
    /// Based on the legacy OutlookTools.emailTypeResponse struct.
    /// </summary>
    public class EmailClassificationResult
    {
        public EmailType Type { get; set; } = EmailType.Unknown;
        public bool IsSubjectLineValid { get; set; } = true;
        public string InvalidReason { get; set; }
        public List<string> PrefIds { get; set; } = new List<string>();
        public List<string> AltIds { get; set; } = new List<string>();
        public List<string> ActivityNums { get; set; } = new List<string>();
        public string Note { get; set; }

        public EmailClassificationResult()
        {
            PrefIds = new List<string>();
            AltIds = new List<string>();
            ActivityNums = new List<string>();
        }
    }
}
