// In IC_Loader_Pro/ViewModels/NoGisFilesViewModel.cs

using ArcGIS.Desktop.Framework;
using ArcGIS.Desktop.Framework.Contracts;
using IC_Loader_Pro.Models;
using IC_Loader_Pro.Services;
using System;
using System.Runtime.InteropServices;
using System.Windows.Input;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace IC_Loader_Pro.ViewModels
{
    public class NoGisFilesViewModel : PropertyChangedBase
    {
        private readonly EmailItem _email;
        private readonly string _fullFolderPath;
        private readonly Outlook.Application _outlookApp;

        public string Sender => _email.SenderEmailAddress;
        public string Subject => _email.Subject;
        public ICommand ViewEmailCommand { get; }

        public NoGisFilesViewModel(EmailItem email, string fullFolderPath, Outlook.Application outlookApp)
        {
            _email = email;
            _fullFolderPath = fullFolderPath;
            _outlookApp = outlookApp;
            ViewEmailCommand = new RelayCommand(OnViewEmail);
        }

        private void OnViewEmail()
        {
            if (_outlookApp == null) return;
            try
            {
                var outlookService = new OutlookService();
                outlookService.DisplayEmailById(_outlookApp, _email.Emailid, _fullFolderPath);
            }
            catch (Exception ex)
            {
                Module1.Log.RecordError("Failed to display email from NoGisFiles dialog.", ex, "OnViewEmail");
                ArcGIS.Desktop.Framework.Dialogs.MessageBox.Show("Could not open the email in Outlook.", "Error");
            }
        }
    }
}