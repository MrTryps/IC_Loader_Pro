using ArcGIS.Desktop.Framework;
using ArcGIS.Desktop.Framework.Contracts;
using IC_Loader_Pro.Models;
using IC_Rules_2025;
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

        public ObservableCollection<string> ToRecipients { get; } = new ObservableCollection<string>();
        public ObservableCollection<string> Attachments { get; } = new ObservableCollection<string>();
        public ICommand ShowResultsCommand { get; }

        public EmailPreviewViewModel(OutgoingEmail email, IcTestResult testResult)
        {
            _emailModel = email;
            Subject = email.Subject;
            HtmlBody = email.Body;

             // It takes the recipients and attachments from the email model
            // and copies them into the collections that the UI is bound to.
            if (email.ToRecipients != null)
            {
                email.ToRecipients.ForEach(r => ToRecipients.Add(r));
            }
            if (email.Attachments != null)
            {
                email.Attachments.ForEach(a => Attachments.Add(a));
            }
            ShowResultsCommand = new RelayCommand(OnShowResults, () => _testResult != null);
        }

        private void OnShowResults()
        {
            // Call the static helper method to show the window
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

            _emailModel.ToRecipients.Clear();
            foreach (var recipient in this.ToRecipients)
            {
                _emailModel.ToRecipients.Add(recipient);
            }

            _emailModel.Attachments.Clear();
            foreach (var attachment in this.Attachments)
            {
                _emailModel.Attachments.Add(attachment);
            }
        }
    }
}