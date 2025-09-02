// In IC_Loader_Pro/Views/EmailPreviewWindow.xaml.cs

using IC_Loader_Pro.ViewModels;
using System.Windows;

namespace IC_Loader_Pro.Views
{
    public partial class EmailPreviewWindow : Window
    {
        private EmailPreviewViewModel ViewModel => DataContext as EmailPreviewViewModel;

        public EmailPreviewWindow()
        {
            InitializeComponent();
            // Subscribe to the Loaded event to set the initial preview
            this.Loaded += EmailPreviewWindow_Loaded;
        }

        private void EmailPreviewWindow_Loaded(object sender, RoutedEventArgs e)
        {
            // Set the initial content of the WebBrowser
            if (ViewModel != null)
            {
                PreviewBrowser.NavigateToString(ViewModel.HtmlBody ?? "");
            }
        }

        private void RefreshButton_Click(object sender, RoutedEventArgs e)
        {
            if (ViewModel != null)
            {
                // Update the preview with the current content of the HTML text box
                PreviewBrowser.NavigateToString(ViewModel.HtmlBody ?? "");
            }
        }

        private void SendButton_Click(object sender, RoutedEventArgs e)
        {
            // Update the underlying email model with any changes before closing
            ViewModel?.UpdateEmailModel();
            this.DialogResult = true;
        }
    }
}