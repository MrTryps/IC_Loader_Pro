using BIS_Tools_DataModels_2025;
using IC_Loader_Pro.Models;
using IC_Loader_Pro.Services;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Input;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace IC_Loader_Pro.ViewModels
{
    public class ManualEmailClassificationViewModel : ArcGIS.Desktop.Framework.Contracts.PropertyChangedBase
    {
        private readonly EmailItem _email;
        private readonly IcGisTypeSetting _icSetting;
        private readonly Outlook.Application _outlookApp;

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

        public ICommand ViewEmailCommand { get; }

        public ManualEmailClassificationViewModel(EmailItem email, IcGisTypeSetting icSetting, Outlook.Application outlookApp)
        {
            _email = email;
            _icSetting = icSetting;
            _outlookApp = outlookApp;

            Sender = email.SenderEmailAddress;
            Subject = string.IsNullOrWhiteSpace(email.Subject) ? "[No Subject]" : email.Subject;
            AttachmentFileNames = email.Attachments.Select(a => a.FileName).ToList();

            AvailableEmailTypes = EmailType.ListProcessableTypes()
                                      .Where(e => e == EmailType.CEA || e == EmailType.DNA || e == EmailType.WRS)
                                      .ToList();
            SelectedEmailType = AvailableEmailTypes.FirstOrDefault();

            ViewEmailCommand = new ArcGIS.Desktop.Framework.RelayCommand(OnViewEmail);
        }

        private void OnViewEmail()
        {
            if (_outlookApp == null || _email == null || _icSetting == null) return;

            var outlookService = new OutlookService();
            outlookService.DisplayEmailById(_outlookApp, _email.Emailid, _icSetting.OutlookInboxFolderPath);
        }
    }
}