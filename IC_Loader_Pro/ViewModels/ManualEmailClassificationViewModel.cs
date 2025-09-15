using IC_Loader_Pro.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using BIS_Tools_DataModels_2025;

namespace IC_Loader_Pro.ViewModels
{
    public class ManualEmailClassificationViewModel : ArcGIS.Desktop.Framework.Contracts.PropertyChangedBase
    {
        public string Sender { get; }
        public string Subject { get; }
        public List<string> AttachmentFileNames { get; }
        public List<EmailType> AvailableEmailTypes { get; }

        private EmailType _selectedEmailType;
        public EmailType SelectedEmailType
        {
            get => _selectedEmailType;
            set => SetProperty(ref _selectedEmailType, value);
        }

        public ManualEmailClassificationViewModel(string sender, string subject, List<string> attachmentNames)
        {
            Sender = sender;
            Subject = string.IsNullOrWhiteSpace(subject) ? "[No Subject]" : subject;
            AttachmentFileNames = attachmentNames;

            // --- START OF THE FIX ---
            // Instead of Enum.GetValues(), we use the static ListProcessableTypes() method from our smart enum class.
            // This is the correct way to get the values from your custom EmailType class.
            AvailableEmailTypes = EmailType.ListProcessableTypes()
                                           .Where(e => e == EmailType.CEA || e == EmailType.DNA || e == EmailType.WRS)
                                           .ToList();
            // --- END OF THE FIX ---

            SelectedEmailType = AvailableEmailTypes.FirstOrDefault();
        }
    }
}