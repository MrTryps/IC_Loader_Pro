using ArcGIS.Desktop.Framework.Contracts;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IC_Loader_Pro.Models
{
    public class EmailItem : PropertyChangedBase
    {
        /// <summary>
        /// The permanent, unchanging Internet Message ID of the email.
        /// Stored in the database's "emailid" column.
        /// </summary>
        private string _emailid;
        public string Emailid
        {
            get => _emailid;
            set => SetProperty(ref _emailid, value);
        }

        private string _subject;
        public string Subject
        {
            get => _subject;
            set => SetProperty(ref _subject, value);
        }

        private DateTime _receivedTime;
        public DateTime ReceivedTime
        {
            get => _receivedTime;
            set => SetProperty(ref _receivedTime, value);
        }

        private string _senderName;
        public string SenderName
        {
            get => _senderName;
            set => SetProperty(ref _senderName, value);
        }

        private string _senderEmailAddress;
        public string SenderEmailAddress
        {
            get => _senderEmailAddress;
            set => SetProperty(ref _senderEmailAddress, value);
        }

        private int _attachmentCount;
        public int AttachmentCount
        {
            get => _attachmentCount;
            set => SetProperty(ref _attachmentCount, value);
        }

        private string _PrefID;
        public string PrefID
        {
            get => _PrefID;
            set => SetProperty(ref _PrefID, value);
        }

        private string _DelID;
        public string DelID
        {
            get => _DelID;
            set => SetProperty(ref _DelID, value);
        }
    }
}
