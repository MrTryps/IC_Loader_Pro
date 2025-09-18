// In IC_Loader_Pro/Views/NoGisFilesWindow.xaml.cs

using System.Windows;

namespace IC_Loader_Pro.Views
{
    public partial class NoGisFilesWindow : Window
    {
        public UserChoice Result { get; private set; }

        public enum UserChoice
        {
            Correspondence,
            Fail,
            Cancel
        }

        public NoGisFilesWindow()
        {
            InitializeComponent();
            this.Result = UserChoice.Cancel;
        }

        private void CorrespondenceButton_Click(object sender, RoutedEventArgs e)
        {
            this.Result = UserChoice.Correspondence;
            this.DialogResult = true;
        }

        private void FailButton_Click(object sender, RoutedEventArgs e)
        {
            this.Result = UserChoice.Fail;
            this.DialogResult = false;
        }
    }
}