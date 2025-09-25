using System.Windows;
using static IC_Loader_Pro.Views.NoGisFilesWindow;

namespace IC_Loader_Pro.Views
{
    public partial class ManualEmailClassificationWindow : Window
    {
        public UserChoice Result { get; private set; }
        public enum UserChoice
        {
            Classify,
            Junk,
            Cancel
        }


        public ManualEmailClassificationWindow()
        {
            InitializeComponent();
            this.Result = UserChoice.Cancel;
        }

        private void OkButton_Click(object sender, RoutedEventArgs e)
        {
            // This sets the window's result to true so the calling code knows the user clicked "OK"
            this.Result = UserChoice.Classify;
            this.DialogResult = true;
        }

        private void JunkButton_Click(object sender, RoutedEventArgs e)
        {
            this.Result = UserChoice.Junk;
            // We also set DialogResult to true to close the window,
            // but our custom Result property will tell us what to do next.
            this.DialogResult = true;
        }
    }
}