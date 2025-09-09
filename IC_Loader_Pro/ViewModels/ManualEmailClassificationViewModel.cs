using IC_Loader_Pro.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using BIS_Tools_DataModels_2025;

// Note: Ensure your file is named ManualEmailClassificationViewModel.cs
namespace IC_Loader_Pro.ViewModels
{
    // The class is renamed
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

        // The constructor is updated to accept the list of attachment names
        public ManualEmailClassificationViewModel(string sender, string subject, List<string> attachmentNames)
        {
            Sender = sender;
            Subject = string.IsNullOrWhiteSpace(subject) ? "[No Subject]" : subject;
            AttachmentFileNames = attachmentNames; // Set the property

            AvailableEmailTypes = Enum.GetValues(typeof(BIS_Tools_DataModels_2025.EmailType))
                                      .Cast<EmailType>()
                                      .Where(e => e == EmailType.CEA || e == EmailType.DNA || e == EmailType.WRS)
                                      .ToList();
            SelectedEmailType = AvailableEmailTypes.FirstOrDefault();
        }
    }
}