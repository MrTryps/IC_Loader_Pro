using ArcGIS.Desktop.Framework;
using ArcGIS.Desktop.Framework.Contracts;
using IC_Loader_Pro.Models;
using IC_Rules_2025;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows.Input;

namespace IC_Loader_Pro.ViewModels
{
    public class EmailPreviewViewModel : PropertyChangedBase
    {
        private OutgoingEmail _emailModel;
        private readonly IcTestResult _testResult;

        private string _subject;
        public string Subject { get => _subject; set => SetProperty(ref _subject, value); }

        private string _htmlBody;
        public string HtmlBody { get => _htmlBody; set => SetProperty(ref _htmlBody, value); }

        // --- START OF MODIFIED CODE ---
        private List<string> _toRecipients;
        public List<string> ToRecipients { get => _toRecipients; set => SetProperty(ref _toRecipients, value); }

        private List<string> _ccRecipients;
        public List<string> CcRecipients { get => _ccRecipients; set => SetProperty(ref _ccRecipients, value); }

        private List<string> _bccRecipients;
        public List<string> BccRecipients { get => _bccRecipients; set => SetProperty(ref _bccRecipients, value); }
        // --- END OF MODIFIED CODE ---

        public ObservableCollection<string> Attachments { get; } = new ObservableCollection<string>();
        public ICommand ShowResultsCommand { get; }

        public EmailPreviewViewModel(OutgoingEmail email, IcTestResult testResult)
        {
            _emailModel = email;
            _testResult = testResult;
            Subject = email.Subject;
            HtmlBody = email.Body;

            // Initialize the lists from the model
            ToRecipients = new List<string>(email.ToRecipients);
            CcRecipients = new List<string>(email.CcRecipients);
            BccRecipients = new List<string>(email.BccRecipients);

            if (email.Attachments != null)
            {
                email.Attachments.ForEach(a => Attachments.Add(a));
            }
            ShowResultsCommand = new RelayCommand(OnShowResults, () => _testResult != null);
        }

        private void OnShowResults()
        {
            Dockpane_IC_LoaderViewModel.ShowTestResultWindow(_testResult);
        }

        /// <summary>
        /// Updates the original email model with any edits made in the UI.
        /// </summary>
        public void UpdateEmailModel()
        {
            _emailModel.Subject = this.Subject;

            _emailModel.OpeningText.Clear();
            _emailModel.MainBodyText.Clear();
            _emailModel.ClosingText.Clear();
            _emailModel.MainBodyText.Add(this.HtmlBody);

            // --- START OF MODIFIED CODE ---
            _emailModel.ToRecipients.Clear();
            _emailModel.ToRecipients.AddRange(this.ToRecipients);

            _emailModel.CcRecipients.Clear();
            _emailModel.CcRecipients.AddRange(this.CcRecipients);

            _emailModel.BccRecipients.Clear();
            _emailModel.BccRecipients.AddRange(this.BccRecipients);
            // --- END OF MODIFIED CODE ---

            _emailModel.Attachments.Clear();
            foreach (var attachment in this.Attachments)
            {
                _emailModel.Attachments.Add(attachment);
            }
        }
    }
}