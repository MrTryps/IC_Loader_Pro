using ArcGIS.Desktop.Catalog;
using ArcGIS.Desktop.Framework;
using ArcGIS.Desktop.Framework.Contracts;
using BIS_Tools_DataModels_2025;
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

        public List<EmailTemplate> InsertableTemplates { get; }

        private List<string> _toRecipients;
        public List<string> ToRecipients { get => _toRecipients; set => SetProperty(ref _toRecipients, value); }

        private List<string> _ccRecipients;
        public List<string> CcRecipients { get => _ccRecipients; set => SetProperty(ref _ccRecipients, value); }

        private List<string> _bccRecipients;
        public List<string> BccRecipients { get => _bccRecipients; set => SetProperty(ref _bccRecipients, value); }

        public ObservableCollection<string> Attachments { get; } = new ObservableCollection<string>();

        private string _selectedAttachment;
        public string SelectedAttachment
        {
            get => _selectedAttachment;
            set
            {
                // This SetProperty call is crucial. It notifies the UI that a property has changed.
                SetProperty(ref _selectedAttachment, value);

                // This line tells the RelayCommand to re-evaluate its CanExecute condition.
                (RemoveAttachmentCommand as RelayCommand)?.RaiseCanExecuteChanged();
            }
        }

        public ICommand ShowResultsCommand { get; }
        public ICommand AddAttachmentCommand { get; }
        public ICommand RemoveAttachmentCommand { get; }

        public EmailPreviewViewModel(OutgoingEmail email, IcTestResult testResult, IcNamedTests namedTests)
        {
            _emailModel = email;
            _testResult = testResult;
            Subject = email.Subject;
            HtmlBody = email.Body;

            InsertableTemplates = namedTests.ReturnInsertTemplates();

            // Initialize the lists from the model
            ToRecipients = new List<string>(email.ToRecipients);
            CcRecipients = new List<string>(email.CcRecipients);
            BccRecipients = new List<string>(email.BccRecipients);

            if (email.Attachments != null)
            {
                email.Attachments.ForEach(a => Attachments.Add(a));
            }
            ShowResultsCommand = new RelayCommand(OnShowResults, () => _testResult != null);
            AddAttachmentCommand = new RelayCommand(OnAddAttachment);
            RemoveAttachmentCommand = new RelayCommand(OnRemoveAttachment, () => !string.IsNullOrEmpty(SelectedAttachment));
        }

        /// <summary>
        /// Inserts a given string of text into the HtmlBody at a specific position.
        /// </summary>
        /// <param name="textToInsert">The template text to insert.</param>
        /// <param name="position">The character index (caret position) where the text should be inserted.</param>
        public void InsertTemplateText(string textToInsert, int position)
        {
            if (position < 0 || position > (HtmlBody?.Length ?? 0))
            {
                position = HtmlBody?.Length ?? 0;
            }
            HtmlBody = HtmlBody.Insert(position, textToInsert);
        }

        private void OnAddAttachment()
        {
            var openDialog = new OpenItemDialog
            {
                Title = "Add Attachment",
                MultiSelect = true, // Allow selecting multiple files
            };

            if (openDialog.ShowDialog() == true)
            {
                foreach (var item in openDialog.Items)
                {
                    Attachments.Add(item.Path);
                }
            }
        }

        private void OnRemoveAttachment()
        {
            if (!string.IsNullOrEmpty(SelectedAttachment))
            {
                Attachments.Remove(SelectedAttachment);
            }
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