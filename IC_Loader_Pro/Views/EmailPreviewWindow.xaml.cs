// In IC_Loader_Pro/Views/EmailPreviewWindow.xaml.cs

using IC_Loader_Pro.ViewModels;
using System.Linq;
using System.Windows;
using System.Windows.Controls;

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
            HtmlBodyTextBox.ContextMenu.Opened += ContextMenu_Opened;
        }

        private void EmailPreviewWindow_Loaded(object sender, RoutedEventArgs e)
        {
            // Set the initial content of the WebBrowser
            if (ViewModel != null)
            {
                PreviewBrowser.NavigateToString(ViewModel.HtmlBody ?? "");
            }
        }

        private void ContextMenu_Opened(object sender, RoutedEventArgs e)
        {
            var contextMenu = sender as ContextMenu;
            if (contextMenu == null || ViewModel == null) return;

            contextMenu.Items.Clear();
            contextMenu.Items.Add(new MenuItem { Header = "Insert Template Text", IsEnabled = false });
            contextMenu.Items.Add(new Separator());

            if (ViewModel.InsertableTemplates.Any())
            {
                foreach (var template in ViewModel.InsertableTemplates)
                {
                    var menuItem = new MenuItem { Header = template.TemplateName };
                    menuItem.Tag = template.ReplacementText;
                    menuItem.ToolTip = template.ReplacementText;
                    menuItem.Click += MenuItem_Click;
                    contextMenu.Items.Add(menuItem);
                }
            }
            else
            {
                contextMenu.Items.Add(new MenuItem { Header = "No insertable templates found", IsEnabled = false });
            }
        }

        private void MenuItem_Click(object sender, RoutedEventArgs e)
        {
            var menuItem = sender as MenuItem;
            if (menuItem?.Tag is string textToInsert)
            {
                int caretPosition = HtmlBodyTextBox.CaretIndex;
                ViewModel.InsertTemplateText(textToInsert, caretPosition);

                HtmlBodyTextBox.Focus();
                HtmlBodyTextBox.CaretIndex = caretPosition + textToInsert.Length;
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